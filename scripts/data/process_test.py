#!/usr/bin/env python3
"""示例：读订单 Excel，按客户汇总 paid 订单，写出 result.xlsx。"""

import argparse
import json
from pathlib import Path

from openpyxl import Workbook, load_workbook


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--in", dest="src", required=True)
    parser.add_argument("--out", dest="dst", required=True)
    args = parser.parse_args()

    src = Path(args.src)
    if not src.exists():
        raise SystemExit(f"输入文件不存在：{src}")

    wb = load_workbook(src, read_only=True, data_only=True)
    ws = wb["orders"]

    rows = ws.iter_rows(values_only=True)
    header = [str(c) for c in next(rows)]
    index = {name: i for i, name in enumerate(header)}

    summary = {}
    skipped = 0

    for row in rows:
        if not row or row[index["order_id"]] is None:
            continue
        if str(row[index["status"]]).strip().lower() != "paid":
            skipped += 1
            continue

        customer = str(row[index["customer"]]).strip().upper()
        qty = float(row[index["qty"]] or 0)
        price = float(row[index["unit_price"]] or 0)
        cost = float(row[index["cost"]] or 0)

        item = summary.setdefault(customer, {"orders": 0, "qty": 0, "revenue": 0.0, "cost": 0.0})
        item["orders"] += 1
        item["qty"] += qty
        item["revenue"] += qty * price
        item["cost"] += qty * cost

    wb.close()

    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = "客户汇总"
    out_ws.append(["customer", "orders", "qty", "revenue", "cost", "profit"])

    total_revenue = 0.0
    total_profit = 0.0
    for customer, item in sorted(summary.items(), key=lambda kv: kv[1]["revenue"], reverse=True):
        profit = item["revenue"] - item["cost"]
        total_revenue += item["revenue"]
        total_profit += profit
        out_ws.append([
            customer,
            item["orders"],
            item["qty"],
            round(item["revenue"], 2),
            round(item["cost"], 2),
            round(profit, 2),
        ])

    dst = Path(args.dst)
    dst.parent.mkdir(parents=True, exist_ok=True)
    out_wb.save(dst)

    print(json.dumps({
        "ok": True,
        "file": str(dst),
        "customers": len(summary),
        "skipped_rows": skipped,
        "revenue": round(total_revenue, 2),
        "profit": round(total_profit, 2),
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
