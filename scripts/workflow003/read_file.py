# -*- coding: utf-8 -*-
"""
60103 计算表 · 极简版
读 3 张源表（订单导出 / 物流运费 / 商品配置），把数据填进「计算表模板」。
运行：python 生成计算表.py
输出：输出/60103计算表_<时间戳>.xlsx
"""
import datetime
import re
import shutil
import zipfile
from copy import copy
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
# ============================ 0. 路径与口径 ============================
S = "源文件"
ORDER = r"./data/订单导出数据20261104182737-623523-652214512317620224.xlsx"
FREIGHT = r"./data/60103物流运费-11.2更.xlsx"
PRODUCT = r"./data/商品配置-11.2更.xlsx"
TEMPLATE = r"./data/60103计算表_20261104_183110.xlsx"
NOW = datetime.datetime.now()
OUT = f"输出/60103计算表_{NOW:%Y%m%d_%H%M%S}.xlsx"
FX = 6.65                                  # 美元汇率：AC = AA / FX
MULT = 1.15                                # 服务费：AA 先 ×1.15
EUR, EUR2USD, EU_TAX = 3, 0.9, 3.33        # 欧盟税三列 T / U / V
EU = {"德国", "法国", "意大利", "西班牙", "荷兰", "比利时", "卢森堡", "奥地利",
      "葡萄牙", "爱尔兰", "丹麦", "瑞典", "芬兰", "希腊", "波兰", "捷克", "匈牙利",
      "罗马尼亚", "保加利亚", "克罗地亚", "斯洛文尼亚", "斯洛伐克", "爱沙尼亚",
      "拉脱维亚", "立陶宛", "马耳他", "塞浦路斯"}       # 英国已脱欧
SUR = {"墨西哥": 2.9, "智利": 7, "卡塔尔": 41, "沙特": 40, "阿联酋": 2}
FILL = {"blocked": "FFFFC7CE", "eu": "FFDDEBF7", "remote": "FFE4DFEC", "surcharge": "FFFCE4D6"}
# ============================ 1. 读三张源表 ============================
def load_products():     
    """商品配置 → {SKU前缀: [采购价, 偏远采购价, 国内运费, 偏远国内运费, 重量]}"""
    ws = openpyxl.load_workbook(PRODUCT, data_only=True)["商品信息"]
    d = {}
    for r in range(4, ws.max_row + 1):
        sku = ws[f"C{r}"].value
        if sku:
            d[str(sku).strip()] = [ws[f"{c}{r}"].value for c in "DEFGH"]
    return d
def seg(text):
    """'0-0.5KG' / '0.001 - 2' / '0.3＜W≤0.5' → (下界, 上界)；空 → None"""
    t = re.sub(r"(?i)kg|\s", "", str(text or ""))
    if "＜W≤" in t:
        a, b = t.split("＜W≤")
    elif "-" in t:
        a, b = t.split("-")
    else:
        return None
    return float(a), float(b)
def load_freight():
    """运价表 → {国家: [(下界, 上界, 公斤价, 挂号费), ...]}"""
    ws = openpyxl.load_workbook(FREIGHT, data_only=True).worksheets[0]
    d, country = {}, None
    for r in range(2, ws.max_row + 1):
        country = ws[f"A{r}"].value or country                    # 国家只在首行，向下补
        s, p, q = seg(ws[f"D{r}"].value), ws[f"E{r}"].value, ws[f"F{r}"].value
        if country and s and isinstance(p, (int, float)) and isinstance(q, (int, float)):
            d.setdefault(str(country).strip(), []).append((s[0], s[1], p, q))
    return d
def load_orders():
    """订单导出 → [{订单信息, items: [(件数, SKU), ...]}]"""
    ws = openpyxl.load_workbook(ORDER, data_only=True).worksheets[0]
    orders, cur = [], None
    for r in range(2, ws.max_row + 1):
        if ws[f"A{r}"].value:                                     # 交易编号非空 = 新订单
            cur = {"txn": str(ws[f"A{r}"].value), "no": ws[f"B{r}"].value,
                   "country": ws[f"E{r}"].value or "", "logi": ws[f"I{r}"].value or "",
                   "chan": ws[f"J{r}"].value or "", "inner": ws[f"K{r}"].value or "",
                   "post": str(ws[f"N{r}"].value or ""), "items": [], "src": r - 1}
            orders.append(cur)
        qty = int(float(ws[f"D{r}"].value or 1))                  # 件数（源表里是文本）
        cur["items"].append((qty, str(ws[f"G{r}"].value or ""), r - 1))
    return orders
# ============================ 2. 匹配与计算 ============================
def prod_of(sku, products):
    """SKU → 商品配置：前缀最长匹配（CHERRY-BEI-* → CHERRY-BEI）"""
    hit = [k for k in products if sku == k or sku.startswith(k + "-")]
    return products[max(hit, key=len)] if hit else None
def rate_of(rates, country, w):
    """重量 → (公斤价, 挂号费)：命中段；更轻按最低段；更重返回 (None, None)"""
    segs = sorted(rates.get(country, []))
    if not segs:
        return None, None
    for lo, hi, p, q in segs:
        if lo < w <= hi:
            return p, q
    return (segs[0][2], segs[0][3]) if w < segs[0][0] else (None, None)
def au_zone(post):
    """澳洲邮编 → 区号（★占位规则，请换成官方分区表）"""
    pc = int(re.sub(r"\D", "", post)[:4] or 0)
    return 1 if 6000 <= pc <= 6999 else 3 if 4000 <= pc <= 4999 else 2
def calc(orders, products, rates):
    """逐单算：重量 → 运价 → 服务费 → 汇率 → 欧盟税 → 账单价格"""
    out, issues = [], []
    for o in orders:
        country, notes = o["country"], []
        if country == "美国" and "岛屿" in o["chan"]:              # 岛屿专线 → 美国（偏远）
            country, notes = "美国（偏远）", ["美国（偏远）专线"]
        elif country == "澳大利亚":
            country = f"澳大利亚（{au_zone(o['post'])}区）"
        remote = country == "美国（偏远）"
        items, weight = [], 0.0
        for qty, sku, src in o["items"]:
            p = prod_of(sku, products)
            if p:
                buy = (p[1] if remote and p[1] else p[0]) or 0
                dom = (p[3] if remote and p[3] else p[2]) or 0
                wt = p[4] or 0.17
            else:
                buy, dom, wt = 15, 4, 0.17                         # 兜底
                issues.append(("警告", o["txn"], o["no"], src, f"SKU「{sku}」未在产品配置中"))
            items.append((qty, sku, buy, dom, wt))
            weight += qty * wt
        weight = round(weight, 3)
        price, fee = rate_of(rates, country, weight)
        blocked = price is None
        if blocked:
            issues.append(("阻断", o["txn"], o["no"], o["src"], f"「{country}」缺 {weight}kg 运价"))
        sur = SUR.get(country, 0)
        tax = EU_TAX if country in EU else 0
        if tax:
            notes.append(f"欧盟国家，收欧盟税 {tax}$")
        if sur:
            notes.append(f"含附加费 {country} +{sur:.1f}￥")
        if country == "美国" and not o["inner"]:
            issues.append(("警告", o["txn"], o["no"], o["src"], "美国单无内部单号，疑似偏远漏处理"))
        rows, total = [], 0.0
        for i, (qty, sku, buy, dom, _wt) in enumerate(items):
            cost = (buy + dom) * qty
            if i == 0:                                             # 运费 + 附加费只进首行
                rmb = ((0 if blocked else price * weight + fee) + cost + sur) * MULT
            else:
                rmb = cost * MULT
            total += rmb / FX
            rows.append((qty, sku, buy, dom, _wt, rmb, rmb / FX))
        out.append(dict(o=o, country=country, weight=weight, price=price, fee=fee,
                        blocked=blocked, sur=sur, tax=tax, rows=rows,
                        total=total, bill=round(total + tax, 2), notes=notes))
    return out, issues
# ============================ 3. 写进计算表 ============================
LEGEND = ["颜色图例：",
          "阻断：缺运价/缺配置等，金额未计（明细格见「未查询」）",
          "警告：需人工确认（原因见异常表）",
          f"欧盟国家：收欧盟税 {EU_TAX}$/单（明细 T/U/V 列）",
          "美国（偏远）：专线改写并套偏远列价格",
          "含附加费国家（墨西哥/智利/卡塔尔/沙特/阿联酋）",
          "金额格「未查询」= 运价表中未查到该运价"]
def fill_rows(ws, rows):
    """覆盖写入第 2 行起；多出的旧行删掉，比模板行数多时套第 2 行样式"""
    old = ws.max_row
    for i, row in enumerate(rows):
        for j, v in enumerate(row, 1):
            c = ws.cell(row=2 + i, column=j, value=v)
            if 2 + i > old:
                c._style = copy(ws.cell(row=2, column=j)._style)
    if old > 1 + len(rows):
        ws.delete_rows(2 + len(rows), old - 1 - len(rows))
def write(result, issues):
    shutil.copy(TEMPLATE, OUT)                       # 以模板为底：样式/列宽/图片都在里面
    wb = openpyxl.load_workbook(OUT)
    ws = wb.worksheets[0]
    ws.title = f"{NOW:%m.%d}计算表"
    cols = [get_column_letter(i) for i in range(1, 39)]
    style = {c: copy(ws[f"{c}2"]._style) for c in cols}          # 记住数据行样式
    height = ws.row_dimensions[2].height
    img = {row[9].value: row[6].value                            # SKU → 图片公式
           for row in ws.iter_rows(min_row=2) if row[9].value and row[6].value}
    for rng in list(ws.merged_cells.ranges):                     # 先拆旧合并
        ws.unmerge_cells(str(rng))
    ws.delete_rows(2, ws.max_row)                                # 清空旧内容，只留表头
    r = 2
    for g in result:
        first, last = r, r + len(g["rows"]) - 1
        for i, (qty, sku, buy, dom, wt, rmb, usd) in enumerate(g["rows"]):
            v = {"A": "", "G": img.get(sku, ""), "H": qty, "J": sku, "K": buy, "L": dom,
                 "M": "万邦普货", "O": "带", "Z": MULT, "AB": FX, "AC": f"=AA{r}/AB{r}"}
            if i == 0:                                            # 订单级字段只写首行
                v.update({"B": g["o"]["logi"], "C": g["o"]["chan"], "D": g["o"]["inner"],
                          "E": g["o"]["no"], "F": g["o"]["txn"], "I": g["country"],
                          "P": g["price"] if not g["blocked"] else "未查询",
                          "Q": g["fee"] if not g["blocked"] else "未查询",
                          "R": "=" + "+".join(f"H{first + k}*{row[4]}"
                                              for k, row in enumerate(g["rows"])),
                          "S": f"=N(P{r})*N(R{r})+N(Q{r})",
                          "AA": f"=(N(S{r})+N(K{r})*N(H{r})+N(L{r})*N(H{r})+N(Y{r}))*N(Z{r})",
                          "AE": f"=SUM(AC{first}:AC{last})",
                          "AD": f"=N(V{r})+AE{r}", "AF": f"=ROUND(AD{r},2)"})
                if g["tax"]:
                    v.update({"T": EUR, "U": EUR2USD, "V": g["tax"], "W": f"=V{r}*AB{r}"})
                if g["sur"]:
                    v["Y"] = g["sur"]
                note = list(g["notes"])
                if g["blocked"]:
                    note.append(f"阻断：「{g['country']}」缺 {g['weight']}kg 运价")
                if note:
                    v["AL"] = "；".join(note)
            else:
                v["AA"] = f"=(N(K{r})*N(H{r})+N(L{r})*N(H{r}))*N(Z{r})"
            for c in cols:
                cell = ws[f"{c}{r}"]
                cell._style = copy(style[c])
                if c in v:
                    cell.value = v[c]
                cell.fill = (PatternFill("solid", fgColor=FILL[cat_of(g)]) if cat_of(g)
                             else PatternFill(fill_type=None))
            ws.row_dimensions[r].height = height
            r += 1
        if last > first:                                          # 国家/重量/账单价格按单合并
            for c in ("I", "R", "AF"):
                ws.merge_cells(f"{c}{first}:{c}{last}")
    ws[f"A{r}"], ws[f"V{r}"] = "总计", f"=SUM(V2:V{r - 1})"
    ws[f"AE{r}"], ws[f"AF{r}"] = f"=SUM(AE2:AE{r - 1})", f"=SUM(AF2:AF{r - 1})"
    for i, text in enumerate(LEGEND):
        ws[f"B{r + 3 + i}"] = text
    # ---- 汇总：按国家小计（阻断单不计）----
    agg = {}
    for g in result:
        if g["blocked"]:
            continue
        a = agg.setdefault(g["country"], dict(o=set(), qty=0, wt=0, intl=0,
                                              dom=0, buy=0, sur=0, eu=0, usd=0))
        a["o"].add(g["o"]["txn"])
        a["qty"] += sum(x[0] for x in g["rows"])
        a["wt"] += g["weight"]
        a["intl"] += g["price"] * g["weight"] + g["fee"]
        a["sur"] += g["sur"]
        a["eu"] += round(g["tax"] * FX, 2)
        a["usd"] += g["bill"]
        for qty, _sku, buy, dom, _wt, _rmb, _usd in g["rows"]:
            a["buy"] += buy * qty
            a["dom"] += dom * qty
    rows = [[k, len(a["o"]), a["qty"], round(a["wt"], 2), round(a["intl"], 2),
             round(a["dom"], 2), round(a["buy"], 2), round(a["sur"], 2),
             round(a["eu"], 2), round(a["usd"], 2)] for k, a in agg.items()]
    rows.append(["合计"] + [round(sum(x[j] for x in rows), 2) for j in range(1, 10)])
    fill_rows(wb["汇总"], rows)
    # ---- 异常：阻断在前，警告在后 ----
    fill_rows(wb["异常"], [[lv, txn, no, src, "", msg]
                           for lv, txn, no, src, msg in
                           sorted(issues, key=lambda x: x[0] != "阻断")])
    # ---- 配置快照：更新本次跑批信息 ----
    snap = wb["配置快照"]
    for key, val in {"批次时间": f"{NOW:%Y-%m-%d %H:%M:%S}", "订单源文件": ORDER,
                     "美元汇率": str(FX), "欧盟税费($)": f"{EU_TAX}（25号后5.5条款已忽略）"}.items():
        for i in range(2, snap.max_row + 1):
            if snap[f"A{i}"].value == key:
                snap[f"B{i}"] = val
    wb.save(OUT)
    keep_wps_images(TEMPLATE, OUT)                   # 补回 WPS 单元格图片
    return OUT
def cat_of(g):
    """行底色分类：阻断 > 欧盟 > 偏远 > 附加费"""
    if g["blocked"]:
        return "blocked"
    if g["tax"]:
        return "eu"
    if g["country"] == "美国（偏远）":
        return "remote"
    if g["sur"]:
        return "surcharge"
    return ""
def keep_wps_images(tpl, out):
    """openpyxl 存盘会丢掉 WPS 的 xl/cellimages.xml + media，这里从模板塞回去"""
    with zipfile.ZipFile(tpl) as z:                       # 图片部件从模板取
        parts = {n: z.read(n) for n in z.namelist()
                 if n in ("xl/cellimages.xml", "xl/_rels/cellimages.xml.rels")
                 or n.startswith("xl/media/")}
    with zipfile.ZipFile(out) as z:                       # 目录文件必须用新文件的
        ct = z.read("[Content_Types].xml").decode()
        rels = z.read("xl/_rels/workbook.xml.rels").decode()
    ct = ct.replace("<Default Extension=\"rels\"",
                    '<Default Extension="png" ContentType="image/png"/>'
                    '<Default Extension="rels"', 1)
    if "/xl/cellimages.xml" not in ct:
        ct = ct.replace("</Types>", '<Override PartName="/xl/cellimages.xml" ContentType='
                        '"application/vnd.wps-officedocument.cellimage+xml"/></Types>')
    if "cellImage" not in rels:
        rels = rels.replace("</Relationships>", '<Relationship Id="rId99" Type="http://www.wps.cn'
                            '/officeDocument/2020/cellImage" Target="cellimages.xml"/></Relationships>')
    tmp = out + ".tmp"
    with zipfile.ZipFile(out) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = ct.encode()
            elif item.filename == "xl/_rels/workbook.xml.rels":
                data = rels.encode()
            zout.writestr(item, data)
        have = set(zin.namelist())
        for n, data in parts.items():
            if n not in have:
                zout.writestr(n, data)
    shutil.move(tmp, out)
if __name__ == "__main__":
    result, issues = calc(load_orders(), load_products(), load_freight())
    path = write(result, issues)
    print(f"✔ {path}")
    print(f"  {len(result)} 单 / {sum(len(g['rows']) for g in result)} 行 / "
          f"阻断 {sum(1 for g in result if g['blocked'])} 单")
    for lv, txn, no, _src, msg in issues:
        print(f"  [{lv}] {txn} {no} {msg}")