# -*- coding: utf-8 -*-
"""
把 n8n 算好的明细写成「计算表」（套模板，保留样式/合并/商品图/公式）。

与 generate.py 的区别：这里不重新算钱，只负责排版。
n8n 传进来的 JSON 结构：
    {
      "detail": [ {交易编号, 订单编号, 国家, SKU, 件数, 单件重量, 公里价... }, ... ],
      "missingList": [ {级别, 类型, 标识, 说明, 订单数, 行数, 样例}, ... ],
      "params": { "FX": 6.65, "EU_TAX": 3.33, "MULT": 1.15, ... }
    }

用法：
    python write_from_template.py --detail /data/.../bill_detail.json \
                                  --out /data/.../output/60103计算表_xxx.xlsx
"""
import argparse
import io
import json
import os
import shutil
import sys
import time
import urllib.request
import zipfile
from copy import copy

import openpyxl
from openpyxl.worksheet.properties import PageSetupProperties
import pandas as pd
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import paths  # noqa: E402

IMG_MAX = 60
FILL = {"blocked": "FFFFC7CE", "eu": "FFDDEBF7", "remote": "FFE4DFEC", "surcharge": "FFFCE4D6"}
_thumb_cache = {}

HERO = "M"          # 模板固定列
HERO_VALUE = "万邦普货"
BATTERY = "O"
BATTERY_VALUE = "带"


# ============================================================================
# 商品图
# ============================================================================
IMG_CACHE_DIR = "/data/data/workflow005/cache/imgs"


def _cache_path(url, size):
    import hashlib
    h = hashlib.sha1(f"{url}|{size}".encode()).hexdigest()[:20]
    return os.path.join(IMG_CACHE_DIR, h + ".png")


def _thumb(url, size=IMG_MAX, attempts=4):
    """下载并缩略；优先用磁盘缓存，失败重试 4 次（指数退避）"""
    if url in _thumb_cache:
        return _thumb_cache[url]
    cp = _cache_path(url, size)
    try:
        if os.path.exists(cp) and os.path.getsize(cp) > 200:
            with open(cp, "rb") as f:
                _thumb_cache[url] = f.read()
            return _thumb_cache[url]
    except Exception:
        pass
    last = None
    for i in range(attempts):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=25) as resp:
                raw = resp.read()
            from PIL import Image
            im = Image.open(io.BytesIO(raw)).convert("RGB")
            im.thumbnail((size * 2, size * 2))
            buf = io.BytesIO()
            im.save(buf, format="PNG", optimize=True)
            _thumb_cache[url] = buf.getvalue()
            try:
                os.makedirs(IMG_CACHE_DIR, exist_ok=True)
                with open(cp, "wb") as f:
                    f.write(_thumb_cache[url])
            except Exception:
                pass
            return _thumb_cache[url]
        except Exception as e:
            last = e
            time.sleep(1.5 * (i + 1))
    raise last


def prime_images(urls, workers=8):
    """并发预下载所有图片（只下载，不插入），失败的交给 insert_images 重试"""
    from concurrent.futures import ThreadPoolExecutor
    todo = [u for u in dict.fromkeys(urls) if isinstance(u, str) and u.startswith("http")
            and "SHIP-PROTECT" not in u.upper() and u not in _thumb_cache]
    if not todo:
        return 0, 0
    ok = fail = 0

    def one(u):
        nonlocal ok, fail
        try:
            _thumb(u)
            ok += 1
        except Exception:
            fail += 1

    with ThreadPoolExecutor(max_workers=workers) as ex:
        list(ex.map(one, todo))
    return ok, fail


def fit_size(w, h, box=IMG_MAX):
    scale = box / max(w, h)
    return max(1, round(w * scale)), max(1, round(h * scale))


def insert_images(ws, records, col="G"):
    ok = fail = 0
    for r, url in records:
        if not isinstance(url, str) or not url.startswith("http"):
            continue
        if "SHIP-PROTECT" in url.upper():
            continue
        try:
            img = XLImage(io.BytesIO(_thumb(url)))
            img.width, img.height = fit_size(img.width, img.height)
            ws.add_image(img, f"{col}{r}")
            ok += 1
        except Exception:
            fail += 1
    return ok, fail


# ============================================================================
# 底色
# ============================================================================
def cat_of(h):
    if h.get("缺运价") or h.get("缺配置"):
        return "blocked"
    if h.get("欧盟税"):
        return "eu"
    if str(h.get("国家", "")).startswith("美国（偏远"):
        return "remote"
    if h.get("附加费"):
        return "surcharge"
    return ""


# ============================================================================
# 写明细
# ============================================================================
def write_detail(ws, detail, fx, params=None):
    recs = {}
    for d in detail:
        recs.setdefault(str(d["交易编号"]), []).append(d)

    cols = [get_column_letter(i) for i in range(1, 39)]
    style = {c: copy(ws[f"{c}2"]._style) for c in cols}
    height = ws.row_dimensions[2].height

    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))
    ws.delete_rows(2, ws.max_row)

    img_rows = []
    r = 2
    for txn, g in recs.items():
        first, last = r, r + len(g) - 1
        fill = cat_of(g[0])
        for i, row in enumerate(g):
            sku = str(row.get("SKU") or "")
            if sku.upper().startswith("SHIP-PROTECT") or row.get("非实物"):
                img_rows.append((r, None))
            else:
                img_rows.append((r, row.get("图片")))

            fixed = row.get("固定金额合计") or 0
            is_fixed = bool(fixed) or (
                row.get("账单固定金额") not in (None, "") and row.get("账单固定金额") is not None
            )
            packs = int(row.get("国内运费份数") or 0)
            wt = row.get("单件重量")
            wt = 0.17 if wt in (None, "") else wt

            v = {
                "G": "", "H": int(row.get("件数") or 1), "J": row.get("SKU"),
                "K": row.get("采购价"), "L": row.get("国内运费单价"),
                HERO: HERO_VALUE, BATTERY: BATTERY_VALUE,
                "Z": row.get("服务费"), "AB": fx,
                "AC": f"=AA{r}/AB{r}",
            }
            # 单价 / 挂号费
            v["P"] = row.get("公斤价") if not row.get("缺运价") else "未查询"
            if row.get("运费行"):
                # Q = 整单挂号费合计（可能含多组的挂号费）
                v["Q"] = (row.get("挂号费合计") if row.get("挂号费合计") is not None
                          else row.get("挂号费")) if not row.get("缺运价") else "未查询"
            # 重量 / 运费：整单模式只在「运费行」写合并公式，其余行留空
            if row.get("整单模式"):
                if row.get("运费行"):
                    # 重量包含固定金额商品（它们的运费列按运价覆盖单独算）
                    parts = [
                        f"H{r + i}*{x.get('单件重量') if x.get('单件重量') is not None else 0.17}"
                        for i, x in enumerate(g)
                        if not x.get("非实物")
                    ]
                    v["R"] = "=" + ("+".join(parts) if parts else "0")
                    if is_fixed:
                        # 固定金额单：账单按固定价，运费列照常写出来（覆盖价 × 重量 + 挂号费）
                        v["S"] = f"=N(P{r})*N(R{r})+N(Q{r})"
                    else:
                        diff = row.get("覆盖差价") or 0
                        if abs(diff) > 0.001:
                            # 运费 = 段价×整单重量 + 覆盖差价 + 挂号费
                            v["S"] = f"=N(P{r})*N(R{r})+({diff:g})+N(Q{r})"
                        else:
                            v["S"] = f"=N(P{r})*N(R{r})+N(Q{r})"
            else:
                v["R"] = f"=H{r}*{wt}"
                if not is_fixed:
                    v["S"] = (f"=N(P{r})*N(R{r})+N(Q{r})" if row.get("运费行")
                              else f"=N(P{r})*N(R{r})")

            if i == 0:
                # AL 备注（和老版一致）：欧盟税 / 附加费 / 美国偏远
                notes = []
                if row.get("欧盟税"):
                    notes.append("欧盟国家，收欧盟税 3.33$")
                if row.get("附加费"):
                    notes.append(f"含附加费 {row.get('国家')} +{row['附加费']:g}￥")
                if str(row.get("国家", "")).startswith("美国（偏远"):
                    notes.append("美国（偏远）专线")
                if row.get("缺运价"):
                    notes.append(f"阻断：{row.get('国家')} 缺 {row.get('整单重量')}kg 运价")
                if row.get("缺配置"):
                    notes.append("阻断：商品配置缺数据")
                if notes:
                    v["AL"] = "；".join(notes)

                v.update({
                    "B": row.get("物流公司"), "C": row.get("物流渠道"),
                    "D": row.get("内部单号"), "E": row.get("订单编号"),
                    "F": row.get("交易编号"), "I": row.get("国家"),
                    "AE": f"=SUM(AC{first}:AC{last})",
                    "AD": f"=N(V{r})+AE{r}",
                    "AF": f"=ROUND(AD{r},2)",
                })
                if row.get("欧盟税"):
                    v.update({"T": 3, "U": 0.9, "V": row["欧盟税"], "W": f"=V{r}*AB{r}"})
                if row.get("附加费"):
                    v["Y"] = row["附加费"]

                # AL 备注（和老版一致）：欧盟税 / 附加费 / 美国偏远
                notes = []
                if row.get("欧盟税"):
                    try:
                        notes.append("欧盟国家，收欧盟税 %g$" % float(row["欧盟税"]))
                    except Exception:
                        notes.append("欧盟国家，收欧盟税")
                if row.get("附加费"):
                    notes.append("含附加费 %s +%g￥" % (row.get("国家") or "", float(row["附加费"])))
                if str(row.get("国家") or "").startswith("美国（偏远"):
                    notes.append("美国（偏远）专线")
                if notes:
                    v["AL"] = "；".join(notes)

            # AA 总价
            terms = []
            if not is_fixed:
                terms.append(f"N(S{r})")
            terms.append(f"N(K{r})*N(H{r})")
            terms.append(f"{packs}*N(L{r})")
            if i == 0 and row.get("附加费"):
                terms.append(f"N(Y{r})")
            v["AA"] = "=(" + "+".join(terms) + f")*N(Z{r})"
            if is_fixed:
                v["AA"] = f"={float(fixed)}*AB{r}"

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
        # 订单备注写 AL（首行）
        note = g[0].get("备注")
        if note:
            ws[f"AL{first}"] = note

    if os.environ.get("BILL_NO_IMAGES"):
        ok, fail = 0, 0
    else:
        _pre_ok, _pre_fail = prime_images([u for _, u in img_rows])
        ok, fail = insert_images(ws, img_rows)

    # 总计块（4 行，对齐客户计算表模板）
    t = r
    ws[f"A{t}"] = "总计"
    try:
        ws[f"A{t}"].font = Font(name=ws[f"A{t}"].font.name, size=ws[f"A{t}"].font.sz, bold=True)
        ws[f"AF{t}"].font = Font(name=ws[f"AF{t}"].font.name, size=ws[f"AF{t}"].font.sz, bold=True)
    except Exception:
        pass
    ws[f"V{t}"] = f"=SUM(V2:V{t - 1})"
    ws[f"AE{t}"] = f"=SUM(AE2:AE{t - 1})"
    ws[f"AF{t}"] = f"=SUM(AF2:AF{t - 1})"

    try:
        eu_tax_lg = float((params or {}).get("EU_TAX_BASE") or (params or {}).get("EU_TAX") or 3.33)
    except Exception:
        eu_tax_lg = 5.6
    # 欧盟补收（可选）：全局参数 EU_TOPUP_COUNT > 0 时才写这一行
    try:
        topup_n = int(float((params or {}).get("EU_TOPUP_COUNT") or 0))
        topup_p = float((params or {}).get("EU_TOPUP_PRICE") or 2.3)
    except Exception:
        topup_n, topup_p = 0, 2.3

    row = t
    if topup_n > 0:
        row = t + 1
        ws[f"AE{row}"] = f"加上{topup_n}个欧盟国家补 {topup_p:g}美元税费"
        ws[f"AF{row}"] = f"={topup_n}*{topup_p:g}"

    fee_row = row + 1
    ws[f"AE{fee_row}"] = 0.005
    ws[f"AF{fee_row}"] = f"=SUM(AF{t}:AF{row})*0.005"

    final_row = fee_row + 1
    ws[f"AF{final_row}"] = f"=SUM(AF{t}:AF{fee_row})"

    # 颜色图例（与客户计算表完全一致的文案与配色）
    ws[f"B{final_row}"] = "颜色图例："
    try:
        ws[f"B{final_row}"].font = copy(ws[f"B{final_row}"].font)
        ws[f"B{final_row}"].font = Font(name=ws[f"B{final_row}"].font.name, size=ws[f"B{final_row}"].font.sz, bold=True)
    except Exception:
        pass
    legend = [
        ("FFFFC7CE", "阻断：缺运价/缺配置等，金额未计（明细格见「未查询」）"),
        ("FFFFEB9C", "警告：需人工确认（如美国单无内部单号、缺创建日期等，原因见行末备注）"),
        ("FFDDEBF7", "欧盟国家：收欧盟税 %g$/单（明细 T/U/V 列）" % eu_tax_lg),
        ("FFE4DFEC", "美国（偏远）：专线改写并套偏远列价格"),
        ("FFFCE4D6", "含附加费国家（墨西哥/智利/卡塔尔/沙特/阿联酋），金额见附加费列"),
        (None, "金额格「未查询」= 运价表中未查到该运价，补录运价表后自动消除"),
    ]
    for i, (color, line) in enumerate(legend, start=1):
        rr = final_row + i
        ws[f"B{rr}"] = line
        if color:
            ws[f"A{rr}"].fill = PatternFill("solid", fgColor=color)

    try:
        eu_tax = float((params or {}).get("EU_TAX_BASE") or (params or {}).get("EU_TAX") or 3.33)
    except Exception:
        eu_tax = 5.6
    legend = [
        "颜色图例：",
        "阻断：缺商品配置/缺运价/缺分区，金额未计（明细格见「未查询」）",
        "警告：需人工确认（原因见异常表）",
        f"欧盟国家：每单收欧盟税 {eu_tax}$（T/U/V 列）",
        "美国（偏远）：专线改写并套偏远列价格",
        "澳大利亚：按邮编查分区，套 1/2/3 区价格",
        "金额格「未查询」= 运价表中未查到该运价",
    ]
    for i, text in enumerate(legend):
        ws[f"B{r + 6 + i}"] = text
    return r, ok, fail


# ============================================================================
# 汇总 / 异常 / 快照
# ============================================================================
SUMMARY_HEADERS = ["国家", "订单数", "件数", "总重量(kg)", "国际运费(￥)", "国内运费(￥)",
                   "采购价(￥)", "附加费(￥)", "欧盟税费(￥)", "账单合计($)"]
ISSUE_HEADERS = ["级别", "类型", "标识", "订单数", "说明", "样例"]
SNAP_HEADERS = ["项目", "值"]


def ensure_sheet(wb, name, headers, widths=None):
    """模板里没有这个 sheet 就新建并写表头"""
    if name in wb.sheetnames:
        return wb[name]
    ws = wb.create_sheet(name)
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = openpyxl.styles.Font(bold=True)
        if widths and i <= len(widths):
            ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    ws.freeze_panes = "A2"
    return ws


def fill_rows(ws, rows, ncol=10):
    old = ws.max_row
    for i, row in enumerate(rows):
        for j, val in enumerate(row, 1):
            c = ws.cell(row=2 + i, column=j, value=val)
            if 2 + i > old:
                c._style = copy(ws.cell(row=2, column=j)._style)
    if old > 1 + len(rows):
        ws.delete_rows(2 + len(rows), old - 1 - len(rows))


def write_summary(wb, detail, fx):
    recs = {}
    for d in detail:
        recs.setdefault(str(d["交易编号"]), []).append(d)

    rows = []
    for txn, g in recs.items():
        if any(x.get("缺运价") or x.get("缺配置") for x in g):
            continue
        h = g[0]
        rows.append([
            h.get("国家"), txn,
            int(sum(int(x.get("件数") or 0) for x in g if not x.get("非实物"))),
            h.get("整单重量"),
            round(sum(float(x.get("运费") or 0) for x in g), 2),
            round(sum(float(x.get("国内运费合计") or 0) for x in g), 2),
            round(sum(int(x.get("件数") or 0) * float(x.get("采购价") or 0) for x in g), 2),
            h.get("附加费") or 0,
            round(float(h.get("欧盟税") or 0) * fx, 2),
            h.get("账单金额") or 0,
        ])
    t = pd.DataFrame(rows, columns=["国家", "交易编号", "件数", "总重量", "国际运费",
                                    "国内运费", "采购价", "附加费", "欧盟税", "账单合计"])
    if t.empty:
        fill_rows(ensure_sheet(wb, "汇总", SUMMARY_HEADERS, [16,8,8,12,14,14,12,12,12,14]),
              [["合计", 0, 0, 0, 0, 0, 0, 0, 0, 0]])
        return
    agg = t.groupby("国家", sort=False).agg(
        订单数=("交易编号", "nunique"), 件数=("件数", "sum"), 总重量=("总重量", "sum"),
        国际运费=("国际运费", "sum"), 国内运费=("国内运费", "sum"),
        采购价=("采购价", "sum"), 附加费=("附加费", "sum"),
        欧盟税=("欧盟税", "sum"), 账单合计=("账单合计", "sum")).round(2).reset_index()
    out = [[a["国家"], a["订单数"], a["件数"], a["总重量"], a["国际运费"], a["国内运费"],
            a["采购价"], a["附加费"], a["欧盟税"], a["账单合计"]] for _, a in agg.iterrows()]
    out.append(["合计"] + [round(sum(x[j] for x in out), 2) for j in range(1, 10)])
    fill_rows(ensure_sheet(wb, "汇总", SUMMARY_HEADERS, [16,8,8,12,14,14,12,12,12,14]), out)


def write_issues(wb, missing_list):
    rows = [[m.get("级别", "阻断"), m.get("类型", ""), m.get("标识", ""), m.get("订单数", 0),
             m.get("说明", ""), m.get("样例", "")] for m in (missing_list or [])]
    if not rows:
        rows = [["", "无", "", 0, "本次全部匹配通过", ""]]
    fill_rows(ensure_sheet(wb, "异常", ISSUE_HEADERS, [8,14,18,8,60,40]), rows, ncol=6)


def write_snapshot(wb, params, detail):
    snap = ensure_sheet(wb, "配置快照", SNAP_HEADERS, [24, 76])
    now = pd.Timestamp.now()
    kv = {
        "批次时间": f"{now:%Y-%m-%d %H:%M:%S}",
        "明细行数": len(detail),
        "订单数": len({d["交易编号"] for d in detail}),
        "美元汇率": params.get("FX"),
        "欧元 转 美元": params.get("EUR2USD"),
        "欧盟税(€/单)": params.get("EUR"),
        "欧盟税($/单)": params.get("EU_TAX"),
        "默认服务费": params.get("MULT"),
        "数据来源": "n8n（飞书配置表 + ERP 订单）",
    }
    rows = [(k, v) for k, v in kv.items() if v is not None]
    fill_rows(snap, rows, ncol=2)


# ============================================================================
# 入口
# ============================================================================
def to_pdf(xlsx_path):
    """用 LibreOffice 把 xlsx 转成 PDF；容器里没有 soffice 就返回 None"""
    import shutil as _shutil
    import subprocess

    exe = (_shutil.which("soffice") or _shutil.which("libreoffice")
           or next((x for x in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",)
                    if os.path.exists(x)), None))
    if not exe:
        print("跳过 PDF：容器里没有 soffice（Dockerfile 需装 libreoffice-writer）")
        return None
    outdir = os.path.dirname(os.path.abspath(xlsx_path))
    cmd = [exe, "--headless", "--norestore", "--convert-to", "pdf",
           "--outdir", outdir, xlsx_path]
    try:
        subprocess.run(cmd, check=True, timeout=600,
                       stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    except Exception as e:
        print(f"PDF 转换失败：{type(e).__name__} {e}")
        return None
    pdf = os.path.splitext(xlsx_path)[0] + ".pdf"
    return pdf if os.path.exists(pdf) else None


def keep_wps_images(tpl, out):
    with zipfile.ZipFile(tpl) as z:
        parts = {n: z.read(n) for n in z.namelist()
                 if n in ("xl/cellimages.xml", "xl/_rels/cellimages.xml.rels")
                 or n.startswith("xl/media/")}
    if not parts:
        return out
    with zipfile.ZipFile(out) as z:
        ct = z.read("[Content_Types].xml").decode()
        rels = z.read("xl/_rels/workbook.xml.rels").decode()
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


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--detail", required=True, help="n8n 传来的明细 JSON")
    ap.add_argument("--out", required=True, help="输出 xlsx")
    ap.add_argument("--template", default=None)
    ap.add_argument("--pdf", action="store_true", help="同时生成 PDF（需要容器里有 soffice）")
    ap.add_argument("--paper", default="A4", help="PDF 纸张：A4 / A3（默认 A4 竖版，对齐客户最新 PDF 样本）")
    ap.add_argument("--fit-width", default=1, help="PDF 缩放成几页宽（默认 1）")
    args = ap.parse_args()

    with open(args.detail, encoding="utf-8") as f:
        payload = json.load(f)
    detail = payload.get("detail") or []
    missing = payload.get("missingList") or []
    params = payload.get("params") or {}
    fx = float(params.get("FX") or 6.65)

    if not detail:
        print("RESULT " + json.dumps({"ok": False, "error": "明细为空"}, ensure_ascii=False))
        return 1

    tpl = args.template or paths.template()
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    shutil.copy(tpl, args.out)

    wb = openpyxl.load_workbook(args.out)
    ws = wb.worksheets[0]
    ws.title = f"{pd.Timestamp.now():%m.%d}计算表"

    r, ok_img, fail_img = write_detail(ws, detail, fx, params)


    # 打印 / PDF 设置：竖版 A4（对齐客户 2026-11-6~11-13 的 PDF 样本）、一页宽、重复表头
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3 if args.paper.upper() == "A3" else ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = max(1, int(args.fit_width))
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:1"
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.25
    ws.page_margins.top = ws.page_margins.bottom = 0.4
    write_summary(wb, detail, fx)
    write_issues(wb, missing)
    write_snapshot(wb, params, detail)

    keep = {ws.title, "汇总", "异常", "配置快照"}
    for other in list(wb.worksheets):
        if other.title not in keep and other.max_row <= 1 and other.max_column <= 1:
            wb.remove(other)

    wb.save(args.out)
    keep_wps_images(tpl, args.out)

    pdf_path = None
    if args.pdf:
        pdf_path = to_pdf(args.out)

    print(f"商品图：成功 {ok_img} 张，失败 {fail_img} 张")
    print("RESULT " + json.dumps({
        "ok": True,
        "out": os.path.abspath(args.out),
        "rows": len(detail),
        "orders": len({d["交易编号"] for d in detail}),
        "issues": len(missing),
        "pdf": pdf_path,
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
