#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
用「已核对/已矫正」的计算表 xlsx 生成 Bill PDF。

流程：
  1) 读 Sheet1（模板格式：38 列）
  2) 按交易编号分组，重算每单：件数 / 重量 / 运费 / 总价￥ / 账单$
  3) 抽出每行的商品图（xlsx 内嵌图片，按行锚点匹配）
  4) 用 bill_pdf 的版式渲染 PDF（页眉/表格/末页条款）

用法：
  python bill_pdf_from_xlsx.py --xlsx <计算表.xlsx> [--out <bill.pdf>] [--date "2026. 11. 19"]
"""
import argparse, datetime, json, os, re, sys, tempfile

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import bill_pdf as B


def num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except Exception:
        return None


def parse_weight(formula, qty_of):
    """R 列公式 =H3*0.17+H4*0.17 → 重量合计"""
    if formula is None:
        return 0.0
    s = str(formula)
    if not s.startswith("="):
        n = num(s)
        return n or 0.0
    total = 0.0
    for row, w in re.findall(r"H(\d+)\*([\d.]+)", s):
        total += (qty_of.get(int(row)) or 0) * float(w)
    if total == 0.0:
        m = re.fullmatch(r"=H(\d+)\*([\d.]+)", s.strip())
        if m:
            total = (qty_of.get(int(m.group(1))) or 0) * float(m.group(2))
    return round(total, 3)


def is_placeholder(fn):
    """过滤马帮的灰色 loading 占位图（近似纯色的小图）"""
    try:
        from PIL import Image as PILImage, ImageStat
        im = PILImage.open(fn).convert("RGB")
        if max(im.size) > 200:
            return False
        st = ImageStat.Stat(im)
        mean = sum(st.mean) / 3
        std = sum(st.stddev) / 3
        return std < 12 and mean > 195
    except Exception:
        return False


def load(path):
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.worksheets[0]

    qty_of, rows = {}, []
    for r in range(2, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a and str(a).strip() == "总计":
            break
        h = num(ws.cell(r, 8).value)
        if h:
            qty_of[r] = h
        rows.append(r)

    # 图片：按锚点行归属
    imgs = {}
    for img in getattr(ws, "_images", []):
        try:
            anchor = img.anchor
            fr = getattr(anchor, "_from", None)
            if fr is None:
                continue
            imgs.setdefault(int(fr.row) + 1, []).append(img)
        except Exception:
            continue

    groups, order = {}, []
    cur = None
    for r in rows:
        t = ws.cell(r, 6).value
        if t not in (None, ""):
            cur = str(t).strip()
            if cur not in groups:
                groups[cur] = []
                order.append(cur)
        if cur is None:
            continue
        groups[cur].append(r)

    out = []
    tmp = tempfile.mkdtemp(prefix="xlsximg_")
    for txn in order:
        rs = groups[txn]
        first = rs[0]
        country = ws.cell(first, 9).value
        name = str(country or "")
        code = B.COUNTRY_CODE.get(name, "")
        country_text = B.COUNTRY_TEXT.get(name) or code or name

        qty = sum(int(num(ws.cell(r, 8).value) or 0) for r in rs)
        uhf = num(ws.cell(first, 22).value) or 0.0     # V 欧盟税（美元）

        total_usd = 0.0
        for r in rs:
            H = num(ws.cell(r, 8).value) or 0
            K = num(ws.cell(r, 11).value) or 0
            L = num(ws.cell(r, 12).value) or 0
            P = num(ws.cell(r, 16).value)
            Q = num(ws.cell(r, 17).value) or 0
            Y = num(ws.cell(r, 25).value) or 0
            Z = num(ws.cell(r, 26).value) or 1
            AB = num(ws.cell(r, 28).value) or 6.6
            wt = parse_weight(ws.cell(r, 18).value, qty_of) if P is not None else 0
            S = (P or 0) * wt + Q
            AA = (S + K * H + L * H + Y) * Z
            total_usd += AA / AB
        amount = round(total_usd + uhf, 2)

        files = []
        for r in rs:
            for img in imgs.get(r, []):
                try:
                    fn = os.path.join(tmp, f"{txn}_{r}_{len(files)}.png")
                    ref = img.ref
                    # 统一转成标准 PNG（xlsx 里可能是 webp/jpeg）
                    try:
                        from PIL import Image as PILImage
                        if hasattr(ref, "seek"):
                            ref.seek(0)
                        im = PILImage.open(ref)
                        if im.mode not in ("RGB", "L"):
                            im = im.convert("RGB")
                        im.save(fn, "PNG")
                    except Exception:
                        if hasattr(ref, "save"):
                            ref.save(fn)
                        else:
                            with open(fn, "wb") as f:
                                f.write(ref.getvalue() if hasattr(ref, "getvalue") else bytes(ref))
                    if os.path.getsize(fn) > 200 and not is_placeholder(fn):
                        files.append(fn)
                except Exception as e:
                    print("图片保存失败", txn, r, type(e).__name__, e, file=sys.stderr)

        out.append({"txn": txn, "qty": qty, "country": country_text,
                    "uhf": round(uhf, 2), "amount": amount, "imgs": files})
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--out", default="")
    ap.add_argument("--date", default="")
    ap.add_argument("--header", default="/data/data/workflow005/config/bill_header.png")
    a = ap.parse_args()

    rows = load(a.xlsx)
    if not rows:
        raise SystemExit("计算表里没有数据行")

    out = a.out or (os.path.splitext(a.xlsx)[0] + ".pdf")
    date_text = a.date or (datetime.datetime.now().strftime("%Y. %-m. %-d"))

    total_amount = round(sum(r["amount"] or 0 for r in rows), 2)
    total_uhf = round(sum(r["uhf"] or 0 for r in rows), 2)

    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4

    total = len(rows)
    pages = max(1, (total + B.ROWS_PER_PAGE - 1) // B.ROWS_PER_PAGE)
    if total > B.ROWS_PER_PAGE:
        last = total - (pages - 1) * B.ROWS_PER_PAGE
        if last > B.MAX_ROWS_LAST:
            pages += 1

    c = canvas.Canvas(out, pagesize=A4)
    c.setTitle("60103 Bill")
    for p in range(pages):
        chunk = rows[p * B.ROWS_PER_PAGE:(p + 1) * B.ROWS_PER_PAGE]
        B.draw_page(c, chunk, p * B.ROWS_PER_PAGE + 1, date_text, a.header,
                    p == pages - 1, total_amount, total_uhf)
        c.showPage()
    c.save()
    print(json.dumps({"ok": True, "out": out, "orders": total, "pages": pages,
                      "total": total_amount, "uhf": total_uhf}, ensure_ascii=False))


if __name__ == "__main__":
    main()
