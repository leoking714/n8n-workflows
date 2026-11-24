#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ROLE: 订单明细 JSON / XLSX 写出工具
ENTRY: 被 import（write_xlsx）
CALLED-BY: write_orders_xlsx.py
DEPS: 无
从 ERP 导出一段时间的订单到 JSON / XLSX（用于对账、比对、离线复算）。

用法：
    python export_orders.py --start "2026-11-09 00:00:00" --end "2026-11-14 23:59:59" \
                            --out /data/data/workflow005/input --format both
"""
import argparse
import json
import os
import urllib.request

ERP = "http://host.docker.internal:8000/api/erp-proxy"
SHOP = ["60103Thecandiceshop"]

FIELDS = [
    ("交易编号", "salesRecordNumber"), ("订单编号", "platformOrderId"),
    ("SKU", "stockSku"), ("商品数量", "quantity"),
    ("国家中文", "countryNameCN"), ("物流公司", "myLogisticsName"),
    ("物流渠道", "myLogisticsChannelName"), ("内部单号", "trackNumber1"),
    ("邮政编码", "postCode"), ("原始图片", "originalPictureUrl"),
]


def fetch_day(start, end):
    out, cursor = [], ""
    for _ in range(50):
        body = json.dumps({"api": "order-get-order-list-new", "data": {
            "createDateStart": start, "createDateEnd": end,
            "maxRows": "500", "cursor": cursor, "shopName": SHOP}}).encode()
        req = urllib.request.Request(ERP, data=body, headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=120) as r:
            d = json.loads(r.read()).get("data") or {}
        if not isinstance(d.get("data"), list):
            break
        out += d["data"]
        if not d.get("hasNext"):
            break
        cursor = d.get("nextCursor", "")
    return out


def to_rows(orders):
    rows = []
    for o in orders:
        for it in (o.get("orderItem") or []):
            rows.append({
                "交易编号": o.get("salesRecordNumber"),
                "订单编号": o.get("platformOrderId"),
                "SKU": it.get("stockSku"),
                "商品数量": it.get("quantity"),
                "国家中文": o.get("countryNameCN"),
                "物流公司": o.get("myLogisticsName"),
                "物流渠道": o.get("myLogisticsChannelName"),
                "内部单号": o.get("trackNumber1"),
                "邮政编码": o.get("postCode"),
                "原始图片": it.get("originalPictureUrl"),
            })
    return rows


def write_xlsx(rows, path, start, end):
    """导出的订单写成带格式的 Excel：表头冻结、自动列宽、加筛选"""
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "订单明细"

    cols = ["交易编号", "订单编号", "SKU", "商品数量", "国家中文", "物流公司",
            "物流渠道", "内部单号", "邮政编码", "原始图片"]
    widths = [12, 18, 42, 10, 12, 26, 40, 22, 12, 60]

    ws.append(cols)
    thin = Side(style="thin", color="FFD0D0D0")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    for i, w in enumerate(widths, 1):
        c = ws.cell(row=1, column=i)
        c.font = Font(bold=True, color="FFFFFFFF")
        c.fill = PatternFill("solid", fgColor="FFB26223")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = bd
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 24

    for r in rows:
        ws.append([r.get(c) for c in cols])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(cols)):
        for c in row:
            c.border = bd
            c.alignment = Alignment(vertical="center",
                                    wrap_text=(c.column == len(cols)))
        row[3].number_format = "0"          # 商品数量

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{ws.max_row}"

    # 顶部说明
    ws2 = wb.create_sheet("说明")
    ws2.append(["订单导出", f"{start} ~ {end}"])
    ws2.append(["明细行数", len(rows)])
    ws2.append(["订单数", len({r.get("交易编号") for r in rows})])
    ws2.append(["列说明", "交易编号=同一订单分组键；SKU=商品编码；商品数量=件数"])
    ws2.column_dimensions["A"].width = 12
    ws2.column_dimensions["B"].width = 60

    wb.save(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--start", required=True)
    ap.add_argument("--end", required=True)
    ap.add_argument("--out", default="/data/data/workflow005/input")
    ap.add_argument("--format", default="xlsx", choices=["json", "xlsx", "both"],
                    help="默认只导 xlsx（好看）；需要给程序吃就加 json")
    args = ap.parse_args()

    days = []
    cur = args.start[:10]
    end_day = args.end[:10]
    from datetime import date, timedelta
    y, m, d = map(int, cur.split("-"))
    e = date(*map(int, end_day.split("-")))
    while True:
        day = date(y, m, d)
        if day > e:
            break
        days.append(day.strftime("%Y-%m-%d"))
        y, m, d = (day + timedelta(days=1)).timetuple()[:3]

    orders = []
    for day in days:
        got = fetch_day(f"{day} 00:00:00", f"{day} 23:59:59")
        print(f"  {day} → {len(got)} 单")
        orders += got

    rows = to_rows(orders)
    os.makedirs(args.out, exist_ok=True)
    stamp = f"{args.start[:10].replace('-','')}_{args.end[:10].replace('-','')}"

    if args.format in ("json", "both"):
        p = os.path.join(args.out, f"orders_{stamp}.json")
        with open(p, "w", encoding="utf-8") as f:
            json.dump(rows, f, ensure_ascii=False, indent=1)
        print("JSON:", p, len(rows), "行")

    if args.format in ("xlsx", "both"):
        p = os.path.join(args.out, f"orders_{stamp}.xlsx")
        write_xlsx(rows, p, args.start, args.end)
        print("XLSX:", p)

    print("RESULT " + json.dumps({"ok": True, "orders": len(orders), "rows": len(rows),
                                  "days": len(days)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
