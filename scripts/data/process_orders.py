#!/usr/bin/env python3
"""读取订单 Excel，计算毛利，输出两个 sheet：明细 + 汇总。"""

import argparse
import json
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

HEADERS = ["order_id", "order_date", "customer", "sku", "qty", "amount", "cost", "margin", "margin_rate"]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--in", dest="src", required=True)
    parser.add_argument("--out", dest="dst", required=True)
    args = parser.parse_args()

    src = Path(args.src)
    dst = Path(args.dst)
    if not src.exists():
        raise SystemExit(f"输入文件不存在：{src}")
    dst.parent.mkdir(parents=True, exist_ok=True)

    wb_in = load_workbook(src, data_only=True)
    ws_in = wb_in[wb_in.sheetnames[0]]

    records = []
    for row in ws_in.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        order_id, order_date, customer, sku, qty, amount, cost = row
        amount = float(amount)
        cost = float(cost)
        margin = round(amount - cost, 2)
        rate = (margin / amount) if amount else 0.0
        records.append(
            {
                "order_id": order_id,
                "order_date": str(order_date),
                "customer": customer,
                "sku": sku,
                "qty": int(qty),
                "amount": amount,
                "cost": cost,
                "margin": margin,
                "margin_rate": rate,
            }
        )

    wb = Workbook()
    ws = wb.active
    ws.title = "明细"
    ws.append(HEADERS)
    for r in records:
        ws.append([r[h] for h in HEADERS])

    header_fill = PatternFill("solid", fgColor="DDEBF7")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
    for row in ws.iter_rows(min_row=2):
        row[5].number_format = "#,##0.00"
        row[6].number_format = "#,##0.00"
        row[7].number_format = "#,##0.00"
        row[8].number_format = "0.00%"
    for idx, width in enumerate([12, 12, 12, 10, 6, 12, 12, 12, 12], start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.freeze_panes = "A2"

    summary: dict[str, dict] = {}
    for r in records:
        item = summary.setdefault(
            r["customer"], {"customer": r["customer"], "orders": 0, "qty": 0, "revenue": 0.0, "profit": 0.0}
        )
        item["orders"] += 1
        item["qty"] += r["qty"]
        item["revenue"] = round(item["revenue"] + r["amount"], 2)
        item["profit"] = round(item["profit"] + r["margin"], 2)

    rows = sorted(summary.values(), key=lambda x: x["profit"], reverse=True)
    for item in rows:
        item["margin_rate"] = round(item["profit"] / item["revenue"], 4) if item["revenue"] else 0.0

    ws2 = wb.create_sheet("汇总")
    ws2.append(["customer", "orders", "qty", "revenue", "profit", "margin_rate"])
    for item in rows:
        ws2.append([item["customer"], item["orders"], item["qty"], item["revenue"], item["profit"], item["margin_rate"]])
    for cell in ws2[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
    for row in ws2.iter_rows(min_row=2):
        row[3].number_format = "#,##0.00"
        row[4].number_format = "#,##0.00"
        row[5].number_format = "0.00%"

    wb.save(dst)

    revenue = round(sum(r["amount"] for r in records), 2)
    profit = round(sum(r["margin"] for r in records), 2)
    print(
        json.dumps(
            {
                "rows": len(records),
                "revenue": revenue,
                "profit": profit,
                "margin_rate": round(profit / revenue, 4) if revenue else 0,
                "top_customer": rows[0]["customer"] if rows else None,
                "file": str(dst),
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
