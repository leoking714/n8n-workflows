# -*- coding: utf-8 -*-
"""
60103 计算表 · 极简版
读 3 张源表（订单导出 / 物流运费 / 商品配置），把数据填进「计算表模板」。

运行：python 生成计算表.py
输出：输出/60103计算表_<时间戳>.xlsx
"""
import datetime
import os
import re
import shutil
import zipfile
from copy import copy

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# ============================ 0. 路径与口径 ============================
# 以「脚本所在目录」为根，这样不管在哪个 cwd 下运行（n8n / 定时任务）都能找到文件
BASE = os.environ.get("BILL_BASE") or os.path.dirname(os.path.abspath(__file__))
S = os.path.join(BASE, "源文件")
ORDER = os.path.join(S, "订单导出数据.xlsx")
FREIGHT = os.path.join(S, "物流运费.xlsx")
PRODUCT = os.path.join(S, "商品配置.xlsx")
TEMPLATE = os.path.join(S, "计算表模板.xlsx")

NOW = datetime.datetime.now()
OUT_DIR = os.environ.get("BILL_OUT") or os.path.join(BASE, "输出")
os.makedirs(OUT_DIR, exist_ok=True)          # ★ 目录不存在就建：shutil.copy 不会自动建目录
OUT = os.path.join(OUT_DIR, f"60103计算表_{NOW:%Y%m%d_%H%M%S}.xlsx")

for _f in (ORDER, FREIGHT, PRODUCT, TEMPLATE):          # 缺文件时报一句人话，而不是 traceback
    if not os.path.exists(_f):
        raise SystemExit(f"找不到输入文件：{_f}")

FX = 6.65                                  # 美元汇率：AC = AA / FX
MULT = 1.15                                # 服务费：AA 先 ×1.15
EUR, EUR2USD, EU_TAX = 3, 0.9, 3.33        # 欧盟税三列 T / U / V
EU = {"德国", "法国", "意大利", "西班牙", "荷兰", "比利时", "卢森堡", "奥地利",
      "葡萄牙", "爱尔兰", "丹麦", "瑞典", "芬兰", "希腊", "波兰", "捷克", "匈牙利",
      "罗马尼亚", "保加利亚", "克罗地亚", "斯洛文尼亚", "斯洛伐克", "爱沙尼亚",
      "拉脱维亚", "立陶宛", "马耳他", "塞浦路斯"}       # 英国已脱欧
SUR = {"墨西哥": 2.9, "智利": 7, "卡塔尔": 41, "沙特": 40, "阿联酋": 2}
FILL = {"blocked": "FFFFC7CE", "eu": "FFDDEBF7", "remote": "FFE4DFEC", "surcharge": "FFFCE4D6"}

# 欧盟税三列自洽提示：金额只认 EU_TAX（V 列），T/U 只是留档
if abs(EUR / EUR2USD - EU_TAX) > 0.01:
    print(f"⚠️ 欧盟税三列不一致：T/U = {EUR}/{EUR2USD} → {EUR / EUR2USD:.2f}，但 V = {EU_TAX}"
          f"（金额按 V 算，T/U 会显示成另一套）")

# ============================ 1. 读三张源表 ============================
def load_products():
    """
    读「商品配置」→ {SKU前缀: [采购价, 偏远采购价, 国内运费, 偏远国内运费, 重量]}。

    输入：源文件/商品配置.xlsx 的「商品信息」表（前 3 行是两层表头，数据在第 4 行之后）
    输出：dict，键是 SKU 颜色前缀（如 CHERRY-BEI），值是 5 个数字
    规则：订单 SKU 用「前缀-」匹配到这里（见 prod_of）；美国（偏远）单取后两列价格
    注意：同一前缀出现多行时后面那行会覆盖前面的（本表里 CHERRY-DOT 两行数值相同）；
          价格/重量为空时保持 None，交给 calc() 兜底（采购15 / 国内4 / 重量0.17）
    """
    ws = openpyxl.load_workbook(PRODUCT, data_only=True)["商品信息"]
    d = {}
    for r in range(4, ws.max_row + 1):
        sku = ws[f"C{r}"].value
        if sku:
            d[str(sku).strip()] = [ws[f"{c}{r}"].value for c in "DEFGH"]
    return d


def seg(text):
    """
    把「重量段」文字拆成 (下界, 上界)。

    输入：运价表 D 列，如 '0-0.5KG' / '0.501-1KG' / '0.001 - 2' / '0.3＜W≤0.5' / '0.34KG'
    输出：(下界, 上界) 两个 float；没有重量段（备注行）返回 None
    规则：区间统一按「下界开、上界闭」使用，即 lo < 重量 <= hi（见 rate_of）
    注意：返回 None 的行会被 load_freight 丢掉，否则“任意重量”会抢走所有命中
    """
    t = re.sub(r"(?i)kg|\s", "", str(text or ""))
    if "＜W≤" in t:
        a, b = t.split("＜W≤")
    elif "-" in t:
        a, b = t.split("-")
    else:
        return None
    return float(a), float(b)


def load_freight():
    """
    读「运价表」→ {国家: [(下界, 上界, 公斤价, 挂号费), ...]}。

    输入：源文件/物流运费.xlsx 的第 1 张表
    规则：国家的合并单元格只在首行出现，用 `or country` 向下补
    注意：三个过滤条件一起生效 —— 有国家 + 有重量段 + 公斤价/挂号费是数字，
          于是合并区空行、备注行（如英国首行 30/16）、'未查询' 文本都会被丢掉
    """
    ws = openpyxl.load_workbook(FREIGHT, data_only=True).worksheets[0]
    d, country = {}, None
    for r in range(2, ws.max_row + 1):
        country = ws[f"A{r}"].value or country                    # 国家只在首行，向下补
        s, p, q = seg(ws[f"D{r}"].value), ws[f"E{r}"].value, ws[f"F{r}"].value
        if country and s and isinstance(p, (int, float)) and isinstance(q, (int, float)):
            d.setdefault(str(country).strip(), []).append((s[0], s[1], p, q))
    return d


def load_orders():
    """
    读「订单导出」→ 一个订单一个 dict，items 里按件记商品。

    输入：源文件/订单导出数据.xlsx 的第 1 张表（一行 = 一件商品，订单字段只在首行）
    输出：[{txn, no, country, logi, chan, inner, post, items:[(件数, SKU, 源表行号)]}]
    规则：交易编号非空 = 开一个新订单；否则就是在给当前订单追加商品
    注意：件数在源表里是文本（"1"），要 int(float(...))；
          源表行号 = Excel 行号 - 1，异常表靠它定位回源表
    """
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
    """
    SKU → 商品配置里对应的那一行（最长前缀匹配）。

    输入：订单 SKU（CHERRY-BEI-L_48597851668619_L_Beige）、load_products() 的结果
    输出：该前缀的 5 个数字；匹配不到返回 None
    规则：sku 等于前缀，或 sku 以「前缀-」开头；多个命中取最长的那条
    注意：分隔符只有 "-"，所以 SHIP-PROTECT_50869900640395_（下划线）匹配不到
          SHIP-PROTECT，会走兜底价 —— 这一点和老表行为一致
    """
    hit = [k for k in products if sku == k or sku.startswith(k + "-")]
    return products[max(hit, key=len)] if hit else None


def rate_of(rates, country, w):
    """
    按整单重量取 (公斤价, 挂号费)。

    输入：运价字典、已改写的国家、整单重量
    输出：(公斤价, 挂号费)；缺价时返回 (None, None)
    规则：① 命中 lo < 重量 <= hi 的那一段
          ② 比最低段还轻 → 按最低段（美国只有 0.501-2KG，0.17kg 也报 75/21）
          ③ 比最高段还重 → 缺价（荷兰 0.51kg 超过 0-0.5KG，就是这个情况）
    """
    segs = sorted(rates.get(country, []))
    if not segs:
        return None, None
    for lo, hi, p, q in segs:
        if lo < w <= hi:
            return p, q
    return (segs[0][2], segs[0][3]) if w < segs[0][0] else (None, None)


def au_zone(post):
    """
    澳洲邮编 → 1 / 2 / 3 区。

    输入：订单的邮政编码（如 '4615'、'T2Z 2X5'）
    输出：区号整数
    注意：★这是占位规则★，只从 2 个样本反推（6175→1区、4615→3区），
          正式使用前必须换成云途官方分区表，否则澳洲单会算错区
    """
    pc = int(re.sub(r"\D", "", post)[:4] or 0)
    return 1 if 6000 <= pc <= 6999 else 3 if 4000 <= pc <= 4999 else 2


def calc(orders, products, rates):
    """
    逐单算钱，产出写表要用的 result 和异常清单 issues。

    输入：load_orders / load_products / load_freight 的结果
    输出：result = 一单一个 dict（国家、重量、运价、每行金额、账单价格、备注）；
          issues = [(级别, 交易编号, 订单号, 源表行号, 说明)]
    流程：改写国家 → 逐件配成本并累计重量 → 查运价 → 附加费/欧盟税 → 逐件算金额
    规则：运费和附加费只进订单首行；1.15 乘在人民币上；欧盟税最后按美元加
    注意：阻断单运费按 0、成本照算，整单不进汇总
    """
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
    """
    把二维列表覆盖写入第 2 行起（用于「汇总」「异常」这种纯值表）。

    输入：工作表、rows（每行是一个列表）
    输出：无（原地改工作表）
    注意：模板行数比数据多时删掉多余旧行；数据比模板多时新行套第 2 行样式，
          否则新行没有边框（这是“看起来不齐”的常见原因）
    """
    old = ws.max_row
    for i, row in enumerate(rows):
        for j, v in enumerate(row, 1):
            c = ws.cell(row=2 + i, column=j, value=v)
            if 2 + i > old:
                c._style = copy(ws.cell(row=2, column=j)._style)
    if old > 1 + len(rows):
        ws.delete_rows(2 + len(rows), old - 1 - len(rows))


def write(result, issues, out=None, now=None):
    """
    按模板写出一份新的计算表，返回文件路径。

    输入：calc() 的 result / issues；out 输出路径（默认用文件顶部的 OUT）
    流程：复制模板 → 记住第 2 行样式和 SKU→图片映射 → 清空旧数据行 →
          逐行写明细（值/公式/合并/底色）→ 总计行 → 图例 → 汇总 → 异常 →
          配置快照 → 存盘 → 补回 WPS 图片部件
    注意：必须“复制模板再改”，不能新建 Workbook —— 否则样式、列宽、合并、
          商品图全部丢失，这也是为什么不直接用 pandas 的 to_excel
    """
    out = out or OUT
    now = now or NOW
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    shutil.copy(TEMPLATE, out)                       # 以模板为底：样式/列宽/图片都在里面
    wb = openpyxl.load_workbook(out)
    ws = wb.worksheets[0]
    ws.title = f"{now:%m.%d}计算表"

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

    wb.save(out)
    keep_wps_images(TEMPLATE, out)                   # 补回 WPS 单元格图片
    return out


def cat_of(g):
    """
    判断一个订单该刷什么底色。

    输入：calc() 里的一个订单 dict
    输出：'blocked' / 'eu' / 'remote' / 'surcharge' / ''（无底色）之一
    规则：优先级 阻断 > 欧盟 > 偏远 > 附加费，一行只刷一种
    """
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
    """
    把 WPS 的单元格图片部件从模板补回新文件（原地改 zip）。

    背景：G 列是 =DISPIMG("ID_xxx",1)，图片不在 drawing 里，而在
          xl/cellimages.xml + xl/media/*.png，用 WPS 私有扩展挂在 workbook.xml.rels 上；
          openpyxl 存盘会把这几样一起丢掉，图片就全空了。
    输入：模板路径、刚写好的文件路径
    输出：无（直接覆盖 out）
    注意：① 图片部件从模板取，但 [Content_Types].xml 和 workbook.xml.rels 必须
            读“新文件”的 —— 模板那份声明了 sharedStrings，新文件没有，弄错会打不开；
          ② 还要补 png 的 Default 声明和 cellImage 关系，否则 WPS 不认图
    """
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
