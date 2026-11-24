#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ROLE: 国家码/文案常量 + 旧 reportlab 引擎
ENTRY: 被 import（只取常量：COUNTRY_CODE / COUNTRY_TEXT）
CALLED-BY: bill_docx.py、bill_pdf_from_xlsx.py（不要用它直接出 PDF，已由 bill_docx 取代）
DEPS: 无

按客户 2026-11 的 Bill PDF 样式生成英文账单（A4 竖版）。

版式（从客户 PDF 反解）：
  页眉图 595.3×144 pt（黄底横幅，含 Bill Date 位）
  日期   x=424.1 基线距顶 62.5，8pt
  表头   基线距顶 155：No./Order number/item name/quantity/country/UHF/Amount
  表格   左 27.1 右 568.1 上 144，表头高 16，行高 36
  数据   首行基线距顶 181，每页 17 行；末页最多 11 行（留条款区）
  末页   Total($) 两行（不含/含 UHF）+ Important Notice 条款

用法：
  python bill_pdf.py --detail detail.json --out bill.pdf [--date "2026. 11. 13"]
"""
import argparse, datetime, json, os, urllib.request

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

W, H = 595.2756, 841.8898

# ---------- 几何 ----------
TBL_LEFT, TBL_RIGHT = 27.1, 568.1
TBL_TOP, HEADER_H, ROW_H = 144.0, 16.0, 36.0
COLS = [27.1, 57.1, 123.1, 339.1, 385.1, 461.1, 515.1, 568.1]
HEADER_BASE_TOP = 155.0
FIRST_BASE_TOP = 181.0
DATE_X, DATE_BASE_TOP = 424.1, 62.5
XH = {"no": 35.7, "order": 63.4, "img": 211.3, "qty": 346.5, "country": 408.4,
      "uhf": 464.1, "uhf2": 472.8, "amount": 526.5}
X = {"no": 37.7, "order": 81.2, "img": 129.1, "qty": 359.9, "country": 417.5,
     "slash": 487.1, "uhf": 464.1}
IMG_W, IMG_H, IMG_STEP = 20.7, 26.0, 24.7
AMOUNT_RIGHT = 564.0
ROWS_PER_PAGE, MAX_ROWS_LAST = 17, 11
YELLOW = (1.0, 0.7529, 0.0)
GRID = 0.749

NOTICE = [
    ("Important Notice:", 33.0),
    ("1. Quote Validity:his quotation is valid for 10 days from the quotation date.", 33.0),
    ("2. Payment Terms:100% prepayment is required before order arrangement.", 33.0),
    ("3. Payment Method Fees:Any transaction fees charged by the payment provider or bank shall be borne by the customer.", 33.0),
    ("Reference:", 40.6),
    ("- Alibaba WorldFirst (Local Bank Transfer): +0.5% (Recommended)", 40.6),
    ("- Payoneer: +0.5%", 40.6),
    ("- PayPal: +5% (Only accepted for sample orders under USD 200)", 40.6),
    ("4. Payment Page:To ensure payment security, please complete payment through our official website:", 33.0),
    ("https://globallyfulfill.com/payment", 40.6),
    ("5. Shipping:This quotation is based on current carrier rates. Final shipping cost shall be confirmed before shipment arrangement.", 33.0),
    ("6. Lead Time:", 33.0),
    ("- Inventory at our warehouse: Orders are normally processed within 24 business hours after confirmation.", 40.6),
    ("- Supplier / Pre-order items: Delivery from supplier to our warehouse usually takes approximately 2-4 business days, followed by fulfillment processing.", 40.6),
    ("7. Transit Time:All transit times shown in this quotation are estimated business days.", 33.0),
    ("For any questions, please feel free to contact our customer service team.", 33.0),
    ("Thank you for your trust and support.", 33.0),
]
NOTICE_TOP, NOTICE_GAP = 615.5, 9.0

COUNTRY_TEXT = {
    "美国": "US",
    "美国（偏远）": "Remote parts of the\nUnited States",
    "澳大利亚（1区）": "Australia Zone 1",
    "澳大利亚（2区）": "Australia Zone 2",
    "澳大利亚（3区）": "Australia Zone 3",
}
COUNTRY_CODE = {
    "美国": "US", "加拿大": "CA", "墨西哥": "MX", "阿联酋": "AE", "英国": "GB",
    "以色列": "IL", "意大利": "IT", "新西兰": "NZ", "巴林": "BH", "西班牙": "ES",
    "荷兰": "NL", "德国": "DE", "卡塔尔": "QA", "哥斯达黎加": "CR", "丹麦": "DK",
    "法国": "FR", "比利时": "BE", "瑞士": "CH", "芬兰": "FI", "爱尔兰": "IE",
    "科威特": "KW", "瑞典": "SE", "沙特": "SA", "挪威": "NO", "奥地利": "AT",
    "波兰": "PL", "捷克": "CZ", "葡萄牙": "PT", "希腊": "GR", "日本": "JP",
    "韩国": "KR", "新加坡": "SG", "马来西亚": "MY", "泰国": "TH", "越南": "VN",
    "印度尼西亚": "ID", "印度": "IN", "土耳其": "TR", "巴西": "BR", "智利": "CL",
    "中国香港": "HK", "中国台湾": "TW", "中国澳门": "MO", "澳大利亚": "AU",
}


def load_rows(path):
    data = json.load(open(path, encoding="utf-8"))
    detail = data.get("detail") or []
    groups, order = {}, []
    for r in detail:
        txn = r.get("交易编号")
        if txn not in groups:
            groups[txn] = []
            order.append(txn)
        groups[txn].append(r)

    rows, blocked = [], []
    for txn in order:
        lines = groups[txn]
        if any(x.get("缺配置") or x.get("缺运价") for x in lines):
            blocked.append(str(txn))
            continue
        first = lines[0]
        name = str(first.get("国家") or "")
        country = (COUNTRY_TEXT.get(name) or first.get("国家二字码")
                   or COUNTRY_CODE.get(name) or name)
        rows.append({
            "txn": txn,
            "qty": sum(int(x.get("件数") or 1) for x in lines),
            "country": country,
            "uhf": round(sum(float(x.get("欧盟税") or 0) for x in lines), 2),
            "amount": round(float(first.get("账单金额") or 0), 2),
            "imgs": [x.get("图片") for x in lines if x.get("图片")],
        })
    if blocked:
        raise SystemExit("存在阻断订单，禁止生成账单 PDF（流程应先拦住）：" + ",".join(blocked[:20]))
    return rows


def fetch_image(url, cache="/tmp/billimg"):
    if not url:
        return None
    if os.path.isabs(str(url)) and os.path.exists(str(url)):
        return str(url)
    os.makedirs(cache, exist_ok=True)
    name = os.path.join(cache, str(abs(hash(url))))
    if os.path.exists(name) and os.path.getsize(name) > 0:
        return name
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=15) as r, open(name, "wb") as f:
            f.write(r.read())
        return name
    except Exception:
        return None


def draw_grid(c, n):
    bottom = TBL_TOP + HEADER_H + n * ROW_H
    c.setFillColorRGB(*YELLOW)
    c.rect(TBL_LEFT, H - TBL_TOP - HEADER_H, TBL_RIGHT - TBL_LEFT, HEADER_H, stroke=0, fill=1)
    c.setFillColorRGB(1, 1, 1)
    c.rect(TBL_LEFT, H - bottom, TBL_RIGHT - TBL_LEFT, n * ROW_H, stroke=0, fill=1)
    c.setStrokeColorRGB(GRID, GRID, GRID)
    c.setLineWidth(0.5)
    c.line(TBL_LEFT, H - TBL_TOP, TBL_RIGHT, H - TBL_TOP)
    for i in range(n + 1):
        y = TBL_TOP + HEADER_H + i * ROW_H
        c.line(TBL_LEFT, H - y, TBL_RIGHT, H - y)
    for x in COLS:
        c.line(x, H - TBL_TOP, x, H - bottom)
    c.setFillColorRGB(0, 0, 0)


def draw_banner(c, header_img):
    if header_img and os.path.exists(header_img):
        c.drawImage(ImageReader(header_img), 0, H - 144, width=W, height=144, mask="auto")


def draw_head_text(c, date_text):
    c.setFont("Helvetica", 8)
    c.drawString(DATE_X, H - DATE_BASE_TOP, date_text)
    c.setFont("Helvetica", 8)
    c.drawString(XH["no"], H - HEADER_BASE_TOP, "No.")
    c.drawString(XH["order"], H - HEADER_BASE_TOP, "Order number")
    c.drawString(XH["img"], H - HEADER_BASE_TOP, "item name")
    c.drawString(XH["qty"], H - HEADER_BASE_TOP, "quantity")
    c.drawString(XH["country"], H - HEADER_BASE_TOP, "country")
    c.drawString(XH["amount"], H - HEADER_BASE_TOP, "Amount")
    c.setFont("Helvetica", 6.5)
    c.drawString(XH["uhf"], H - 151.0, "Union Handling")
    c.drawString(XH["uhf2"], H - 158.2, "Fee (UHF)")


def draw_page(c, page_rows, start_no, date_text, header_img, is_last, sum_amount, sum_uhf):
    draw_banner(c, header_img)
    n = max(len(page_rows), 1)
    draw_grid(c, n)
    draw_head_text(c, date_text)

    no = start_no
    for i, r in enumerate(page_rows):
        y = H - (FIRST_BASE_TOP + i * ROW_H)
        c.setFont("Helvetica", 8)
        c.drawString(X["no"], y, str(no))
        c.drawString(X["order"], y, str(r["txn"]))
        c.drawString(X["qty"], y, str(r["qty"]))

        txt = str(r["country"])
        if "\n" in txt:
            a, b = txt.split("\n", 1)
            c.setFont("Helvetica", 6)
            c.drawString(396.8, y + 2.5, a)
            c.drawString(405.1, y - 7.5, b)
            c.setFont("Helvetica", 8)
        else:
            c.drawString(X["country"], y, txt)

        if r["uhf"] and r["uhf"] > 0:
            c.drawString(X["uhf"], y, "US$%.2f" % r["uhf"])
        else:
            c.setFont("Helvetica", 7)
            c.drawString(X["slash"], y, "/")
            c.setFont("Helvetica", 8)

        c.drawRightString(AMOUNT_RIGHT, y, "US$%.2f" % (r["amount"] or 0))

        for j, url in enumerate((r.get("imgs") or [])[:3]):
            img = fetch_image(url)
            if not img:
                continue
            try:
                c.drawImage(img, X["img"] + j * IMG_STEP, y - 8, width=IMG_W, height=IMG_H,
                            preserveAspectRatio=True, mask="auto")
            except Exception:
                pass
        no += 1

    if is_last:
        c.setFont("Helvetica", 8)
        c.drawString(31.1, H - (556.8 + 6.4), "Total($): US$%.2f" % max(0.0, sum_amount - sum_uhf))
        c.drawString(31.1, H - (592.8 + 6.4), "Total($): US$%.2f" % sum_amount)
        c.setFont("Helvetica", 6)
        for i, (text, x) in enumerate(NOTICE):
            c.drawString(x, H - (NOTICE_TOP + i * NOTICE_GAP + 4.8), text)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--detail", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--date", default=None)
    ap.add_argument("--header", default="/data/data/workflow005/config/bill_header.png")
    ap.add_argument("--limit", type=int, default=0)
    a = ap.parse_args()

    rows = load_rows(a.detail)
    if a.limit:
        rows = rows[: a.limit]

    if a.date:
        date_text = a.date
    else:
        t = datetime.datetime.now()
        date_text = f"{t.year}. {t.month}. {t.day}"

    total_amount = round(sum(r["amount"] or 0 for r in rows), 2)
    total_uhf = round(sum(r["uhf"] or 0 for r in rows), 2)

    total = len(rows)
    pages = max(1, (total + ROWS_PER_PAGE - 1) // ROWS_PER_PAGE)
    if total > ROWS_PER_PAGE:
        last = total - (pages - 1) * ROWS_PER_PAGE
        if last > MAX_ROWS_LAST:
            pages += 1

    c = canvas.Canvas(a.out, pagesize=A4)
    c.setTitle("60103 Bill")
    for p in range(pages):
        chunk = rows[p * ROWS_PER_PAGE:(p + 1) * ROWS_PER_PAGE]
        draw_page(c, chunk, p * ROWS_PER_PAGE + 1, date_text, a.header,
                  p == pages - 1, total_amount, total_uhf)
        c.showPage()
    c.save()
    print(json.dumps({"ok": True, "out": a.out, "orders": total, "pages": pages}, ensure_ascii=False))


if __name__ == "__main__":
    main()
