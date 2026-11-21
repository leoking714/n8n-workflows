# -*- coding: utf-8 -*-
"""
把算好的订单明细 df 写进「计算表」模板。

用法（在你现有脚本末尾加两行）：
    from write_table import write_excel, build_issues
    write_excel(detail, build_issues(detail), "输出/60103计算表.xlsx")

要求 detail 至少包含这些列：
    交易编号 / 订单编号 / 物流公司 / 物流渠道 / 内部单号 / 国家 / SKU /
    商品数量 / 采购价 / 国内运费 / 重量 / 服务费 / 订单重量 /
    公斤价 / 挂号费 / 缺运价 / 附加费 / 欧盟税 / 首行
"""
import io
import os
import shutil
import urllib.request
import zipfile
from copy import copy

import openpyxl
import pandas as pd
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image

# ============================================================================
# 参数 —— 改成你自己的路径/口径
# ============================================================================
BASE = os.environ.get("BILL_BASE") or os.path.dirname(os.path.abspath(__file__))

# 模板：11.3 那份计算表（Sheet1 = 明细，38 列 A~AL）
import paths
TEMPLATE = paths.template()
SHEET_NAME = None      # None = 保留模板里的 sheet 名（如 Sheet1）；填字符串则改名
OUT_DIR = os.environ.get("BILL_OUT") or os.path.join(BASE, "输出")

FX = 6.65                        # 人民币 / 美元
MULT = 1.15                      # 默认服务费（毛利列没填时）
EUR = 3.0                        # 欧盟每单多收 3 欧元
EUR2USD = 0.9                    # 欧元 → 美元
EU_TAX = round(EUR / EUR2USD, 2)  # 3.33

EU = {"德国", "法国", "意大利", "西班牙", "荷兰", "比利时", "卢森堡", "奥地利",
      "葡萄牙", "爱尔兰", "丹麦", "瑞典", "芬兰", "希腊", "波兰", "捷克", "匈牙利",
      "罗马尼亚", "保加利亚", "克罗地亚", "斯洛文尼亚", "斯洛伐克", "爱沙尼亚",
      "拉脱维亚", "立陶宛", "马耳他", "塞浦路斯"}

SUR = {"墨西哥": 2.9, "智利": 7.0, "卡塔尔": 41.0, "沙特": 40.0,
       "阿联酋": 0.0,        # 规则 2 元，当前不收
       "罗马尼亚": 42.0}     # 处理费

FILL = {"blocked": "FFFFC7CE", "eu": "FFDDEBF7",
        "remote": "FFE4DFEC", "surcharge": "FFFCE4D6"}

SERVICE_COL = "毛利"              # 运价表里存服务费率的列名（异常/快照里用）
QTY = "商品数量"                  # 你的件数列名
WT = "重量"                       # 你的单件重量列名

# ---- 商品图 ----
# "url"      = 按「原始图片」列的链接下载、缩小后插到 G 列（Excel/WPS 都能看）
# "dispimg"  = 用模板里的 WPS 单元格图片公式（只有 WPS 能显示）
IMAGE_MODE = "url"
IMG_URL_COL = "原始图片"
IMG_MAX = 60                      # 图片最长边（像素），另一边按原比例算
_thumb_cache = {}


# ============================================================================
# 1.5 商品图：按 URL 下载 → 缩小 → 插到 G 列
# ============================================================================
def _thumb(url, size=IMG_MAX):
    """下载图片并缩成小图（同一个 URL 只下一次），返回 PNG bytes"""
    if url in _thumb_cache:
        return _thumb_cache[url]
    with urllib.request.urlopen(url, timeout=20) as resp:
        raw = resp.read()
    im = Image.open(io.BytesIO(raw)).convert("RGB")
    im.thumbnail((size * 2, size * 2))            # 等比缩略（PIL 会保持原比例）
    buf = io.BytesIO()
    im.save(buf, format="PNG", optimize=True)
    _thumb_cache[url] = buf.getvalue()
    return _thumb_cache[url]


def fit_size(w, h, box=IMG_MAX):
    """按原比例缩放：最长边 = box，另一边跟着比例走（横图宽=box，竖图高=box）"""
    scale = box / max(w, h)
    return max(1, round(w * scale)), max(1, round(h * scale))


def insert_images(ws, records, col="G"):
    """records = [(行号, 图片URL)]；下载后锚定到该行 G 列并缩小"""
    ok = fail = 0
    for r, url in records:
        if not isinstance(url, str) or not url.startswith("http"):
            continue
        try:
            img = XLImage(io.BytesIO(_thumb(url)))
            img.width, img.height = fit_size(img.width, img.height)   # 保持原比例
            ws.add_image(img, f"{col}{r}")
            ok += 1
        except Exception as e:
            fail += 1
            print(f"  ⚠️ 第 {r} 行图片没插上：{type(e).__name__} {e}")
    print(f"  商品图：成功 {ok} 张，失败 {fail} 张")
    return ok, fail


# ============================================================================
# 1.8 表头校验 / 建缺失的 sheet
# ============================================================================
# A~AL 38 列的表头（用 startswith 匹配，兼容「件数（要填数量默认填1）」这种长名）
EXPECT_HEADERS = [
    "编号", "物流公司", "物流渠道", "内部单号", "订单编号", "交易编号", "名称图片",
    "件数", "国家", "SKU", "单个采购价", "单个国内运费", "物流公司", "体积", "是否带电",
    "运费单价", "挂号费", "重量", "运费", "欧盟税费（欧元）", "欧元美元汇率",
    "欧盟税费（美元）", "欧盟税费（人民币）", "附加费1", "附加费2", "服务费",
    "总价（￥）", "美元汇率", "总价（$）", "最终总价", "订单总价", "账单价格",
    "成本", "履行价", "毛利润", "时效", "供应商报", "1688",
]


def check_headers(ws):
    """比对模板表头，返回不匹配的列（只提示，不中断）"""
    bad = []
    for i, want in enumerate(EXPECT_HEADERS, 1):
        got = str(ws.cell(row=1, column=i).value or "")
        if not got.startswith(want):
            bad.append(f"{get_column_letter(i)}: 期望「{want}」实际「{got}」")
    if bad:
        print("  ⚠️ 模板表头与本代码不一致：")
        for b in bad:
            print("     ", b)
    return bad


SUMMARY_HEADERS = ["国家", "订单数", "件数", "总重量(kg)", "国际运费(￥)", "国内运费(￥)",
                   "采购价(￥)", "附加费(￥)", "欧盟税费(￥)", "账单合计($)"]
ISSUE_HEADERS = ["级别", "交易编号", "订单编号", "源表行号", "异常代码", "说明"]
SNAP_HEADERS = ["项目", "值"]


def ensure_sheet(wb, name, headers, widths=None):
    """没有这个 sheet 就新建并写表头（有就直接返回）"""
    if name in wb.sheetnames:
        return wb[name]
    ws = wb.create_sheet(name)
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = PatternFill("solid", fgColor="FFD9D9D9")
        if widths:
            ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    ws.freeze_panes = "A2"
    return ws


# ============================================================================
# 2. 行底色分类
# ============================================================================
def cat_of(g):
    """一个订单该刷什么底色：阻断 > 欧盟 > 偏远 > 附加费"""
    h = g.iloc[0]
    if h["缺运价"]:
        return "blocked"
    if h["欧盟税"]:
        return "eu"
    if h["国家"] == "美国（偏远）":
        return "remote"
    if h["附加费"]:
        return "surcharge"
    return ""


# ============================================================================
# 3. 写明细（核心）
# ============================================================================
def write_excel(df, issues, out_path):
    """
    把 calc 出来的明细写成计算表。

    每行公式（P/Q/R/S/AA）逐行写：
        R  = 件数 × 单件重量                  (行重量)
        S  = P × R                            (固定金额行不写 S)
        S  = P × R + Q                        (只有「运费行」加挂号费)
        AA = (S + K×H + 份数×L + Y) × Z        (Y 只首行；固定金额行 AA = 固定金额合计 × AB)
    """
    if "源表行号" not in df.columns:
        df = df.assign(源表行号=range(1, len(df) + 1))

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    shutil.copy(TEMPLATE, out_path)
    wb = openpyxl.load_workbook(out_path)
    ws = wb.worksheets[0]
    if SHEET_NAME:
        ws.title = SHEET_NAME
    check_headers(ws)

    # issues → 每单备注
    notes_by_txn = {}
    for lv, txn, _no, _src, _code, msg in issues:
        notes_by_txn.setdefault(str(txn), []).append(f"{lv}：{msg}")

    cols = [get_column_letter(i) for i in range(1, 39)]
    style = {c: copy(ws[f"{c}2"]._style) for c in cols}
    height = ws.row_dimensions[2].height
    img = {}
    if IMAGE_MODE == "dispimg":
        img = {r[9].value: r[6].value
               for r in ws.iter_rows(min_row=2) if r[9].value and r[6].value}
    img_rows = []

    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))
    ws.delete_rows(2, ws.max_row)

    r = 2
    for txn, g in df.groupby("交易编号", sort=False):
        first, last = r, r + len(g) - 1
        for i, (_, row) in enumerate(g.iterrows()):
            if IMAGE_MODE == "url":
                img_rows.append((r, row.get(IMG_URL_COL)))

            fixed = row.get("账单固定金额")
            fixed = fixed if pd.notna(fixed) else None
            packs = int(row["国内运费份数"]) if pd.notna(row.get("国内运费份数")) else 0
            wt = row[WT]

            v = {"A": "", "G": img.get(row["SKU"], ""), "H": int(row[QTY]),
                 "J": row["SKU"], "K": row["采购价"], "L": row["国内运费"],
                 "M": "万邦普货", "O": "带", "Z": row["服务费"],
                 "AB": FX, "AC": f"=AA{r}/AB{r}"}

            # 逐行：重量
            v["R"] = f"=H{r}*{0.17 if pd.isna(wt) else wt}"

            # 逐行：单价；挂号费只在运费行
            v["P"] = row["公斤价"] if not row["缺运价"] else "未查询"
            if row.get("运费行"):
                v["Q"] = row["挂号费"] if not row["缺运价"] else "未查询"

            # 逐行：运费
            if fixed is None:
                v["S"] = (f"=N(P{r})*N(R{r})+N(Q{r})" if row.get("运费行")
                          else f"=N(P{r})*N(R{r})")

            if i == 0:                                  # 订单级字段
                v.update({"B": row["物流公司"], "C": row["物流渠道"],
                          "D": row["内部单号"], "E": row["订单编号"],
                          "F": row["交易编号"], "I": row["国家"],
                          "AE": f"=SUM(AC{first}:AC{last})",
                          "AD": f"=N(V{r})+AE{r}",
                          "AF": f"=ROUND(AD{r},2)"})
                if row["欧盟税"]:
                    v.update({"T": EUR, "U": EUR2USD, "V": row["欧盟税"],
                              "W": f"=V{r}*AB{r}"})
                if row["附加费"]:
                    v["Y"] = row["附加费"]

                al = []
                if not row["缺运价"]:
                    if row["欧盟税"]:
                        al.append(f"欧盟国家，收欧盟税 {EU_TAX}$")
                    if row["附加费"]:
                        al.append(f"含附加费 {row['国家']} +{row['附加费']:g}￥")
                    if row["国家"] == "美国（偏远）":
                        al.append("美国（偏远）专线")
                if "非实物" in g.columns:
                    n_ng = int(g["非实物"].sum())
                    if n_ng:
                        al.append(f"含非实物 {n_ng} 行（SHIP 保护费等，不计重量/成本）")
                al += notes_by_txn.get(str(row["交易编号"]), [])
                if al:
                    v["AL"] = "；".join(al)

            # 统一拼 AA
            terms = []
            if fixed is None:
                terms.append(f"N(S{r})")
            terms.append(f"N(K{r})*N(H{r})")
            terms.append(f"{packs}*N(L{r})")
            if i == 0 and row.get("附加费"):
                terms.append(f"N(Y{r})")
            v["AA"] = "=(" + "+".join(terms) + f")*N(Z{r})"

            # 固定金额行：总价锁死
            if fixed is not None:
                v["AA"] = f"={float(row['固定金额合计'])}*AB{r}"

            fill = cat_of(g)
            for c in cols:
                cell = ws[f"{c}{r}"]
                cell._style = copy(style[c])
                if c in v:
                    cell.value = v[c]
                cell.fill = (PatternFill("solid", fgColor=FILL[fill]) if fill
                             else PatternFill(fill_type=None))
            ws.row_dimensions[r].height = height
            r += 1
        if last > first:
            for c in ("I", "AF"):
                ws.merge_cells(f"{c}{first}:{c}{last}")

    if IMAGE_MODE == "url" and img_rows:
        insert_images(ws, img_rows)

    # 总计行 + 图例
    ws[f"A{r}"], ws[f"V{r}"] = "总计", f"=SUM(V2:V{r - 1})"
    ws[f"AE{r}"], ws[f"AF{r}"] = f"=SUM(AE2:AE{r - 1})", f"=SUM(AF2:AF{r - 1})"

    legend = ["颜色图例：",
              "阻断：缺运价/缺配置等，金额未计（明细格见「未查询」）",
              "警告：需人工确认（原因见异常表）",
              f"欧盟国家：收欧盟税 {EU_TAX}$/单（{EUR:g}欧元 ÷ {EUR2USD:g}）",
              "美国（偏远）：专线改写并套偏远列价格",
              "含附加费国家（" + "/".join(f"{k}{v:g}" for k, v in SUR.items() if v > 0) + "）",
              "金额格「未查询」= 运价表中未查到该运价"]
    for i, text in enumerate(legend):
        ws[f"B{r + 3 + i}"] = text

    # 汇总
    rows = []
    for txn, g in df.groupby("交易编号", sort=False):
        h = g.iloc[0]
        if h["缺运价"]:
            continue
        qty = int(g.loc[~g["非实物"], QTY].sum()) if "非实物" in g.columns \
            else int(g[QTY].sum())
        rows.append([h["国家"], txn, qty, h["订单重量"], g["运费"].sum(),
                     g["国内运费合计"].sum(),
                     (g[QTY] * g["采购价"]).sum(),
                     h["附加费"], round(h["欧盟税"] * FX, 2), h["账单金额"]])
    t = pd.DataFrame(rows, columns=["国家", "交易编号", "件数", "总重量", "国际运费",
                                    "国内运费", "采购价", "附加费", "欧盟税", "账单合计"])
    agg = t.groupby("国家", sort=False).agg(
        订单数=("交易编号", "nunique"), 件数=("件数", "sum"), 总重量=("总重量", "sum"),
        国际运费=("国际运费", "sum"), 国内运费=("国内运费", "sum"),
        采购价=("采购价", "sum"), 附加费=("附加费", "sum"),
        欧盟税=("欧盟税", "sum"), 账单合计=("账单合计", "sum")).round(2).reset_index()
    summary = [[a["国家"], a["订单数"], a["件数"], a["总重量"], a["国际运费"],
                a["国内运费"], a["采购价"], a["附加费"], a["欧盟税"], a["账单合计"]]
               for _, a in agg.iterrows()]
    summary.append(["合计"] + [round(sum(x[j] for x in summary), 2) for j in range(1, 10)])
    fill_rows(ensure_sheet(wb, "汇总", SUMMARY_HEADERS,
                           [16, 8, 8, 12, 14, 14, 12, 12, 12, 14]), summary)

    # 异常
    fill_rows(ensure_sheet(wb, "异常", ISSUE_HEADERS, [8, 12, 18, 10, 22, 90]),
              [[lv, txn, no, src, code, msg]
               for lv, txn, no, src, code, msg in
               sorted(issues, key=lambda x: x[0] != "阻断")], ncol=6)

    # 配置快照
    snap = ensure_sheet(wb, "配置快照", SNAP_HEADERS, [24, 76])
    now = pd.Timestamp.now()
    fill_rows(snap, [
        ("批次时间", f"{now:%Y-%m-%d %H:%M:%S}"),
        ("明细 sheet", ws.title),
        ("明细行数", len(df)),
        ("订单数", int(df["交易编号"].nunique())),
        ("阻断单数", int(df.loc[df["首行"], "缺运价"].sum())),
        ("美元汇率", FX),
        ("欧元 → 美元", EUR2USD),
        ("欧盟税(欧元/单)", EUR),
        ("欧盟税($/单)", EU_TAX),
        ("默认服务费", MULT),
        ("商品图方式", "按「原始图片」URL 下载" if IMAGE_MODE == "url" else "模板 DISPIMG"),
        ("模板文件", os.path.basename(TEMPLATE)),
    ])

    keep = {ws.title, "汇总", "异常", "配置快照"}
    for other in list(wb.worksheets):
        if other.title in keep:
            continue
        if other.max_row <= 1 and other.max_column <= 1:
            wb.remove(other)

    wb.save(out_path)
    if IMAGE_MODE == "dispimg":
        keep_wps_images(TEMPLATE, out_path)
    return out_path


# ============================================================================
# 4. 纯值表覆盖写入（汇总 / 异常 用）
# ============================================================================
def fill_rows(ws, rows, ncol=10):
    """从第 2 行起覆盖写入；数据比模板少就删掉多余行，多就套第 2 行样式"""
    old = ws.max_row
    for i, row in enumerate(rows):
        for j, val in enumerate(row, 1):
            c = ws.cell(row=2 + i, column=j, value=val)
            if 2 + i > old:
                c._style = copy(ws.cell(row=2, column=j)._style)
    if old > 1 + len(rows):
        ws.delete_rows(2 + len(rows), old - 1 - len(rows))


# ============================================================================
# 5. 把 WPS 单元格图片补回新文件
# ============================================================================
def keep_wps_images(tpl, out):
    """openpyxl 存盘会丢掉 xl/cellimages.xml + xl/media/*，这里从模板塞回去"""
    with zipfile.ZipFile(tpl) as z:
        parts = {n: z.read(n) for n in z.namelist()
                 if n in ("xl/cellimages.xml", "xl/_rels/cellimages.xml.rels")
                 or n.startswith("xl/media/")}
    with zipfile.ZipFile(out) as z:
        ct = z.read("[Content_Types].xml").decode()
        rels = z.read("xl/_rels/workbook.xml.rels").decode()

    ct = ct.replace('<Default Extension="rels"',
                    '<Default Extension="png" ContentType="image/png"/>'
                    '<Default Extension="rels"', 1)
    if "/xl/cellimages.xml" not in ct:
        ct = ct.replace("</Types>", '<Override PartName="/xl/cellimages.xml" '
                        'ContentType="application/vnd.wps-officedocument.cellimage+xml"/></Types>')
    if "cellImage" not in rels:
        rels = rels.replace("</Relationships>", '<Relationship Id="rId99" '
                            'Type="http://www.wps.cn/officeDocument/2020/cellImage" '
                            'Target="cellimages.xml"/></Relationships>')

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
    return out


# ============================================================================
# 6. 自测：没有 df 时直接提示怎么用
# ============================================================================
if __name__ == "__main__":
    raise SystemExit(
        "这是写表模块，不能单独运行。\n"
        "在你现有脚本末尾加两行：\n\n"
        "    from write_table import write_excel, build_issues\n"
        "    write_excel(detail, build_issues(detail), '输出/60103计算表.xlsx')\n"
    )
