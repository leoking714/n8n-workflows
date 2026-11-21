#!/usr/bin/env python3
"""生成测试用 Excel：30~N 行订单，字段固定。"""

import argparse
import json
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook

CUSTOMERS = ["ACME", "NORDIC", "SAKURA", "LUMEN"]
SKUS = ["SKU-A", "SKU-B", "SKU-C"]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", required=True, help="输出 xlsx 路径，如 /data/lab/demo/orders.xlsx")
    parser.add_argument("--rows", type=int, default=40)
    parser.add_argument("--seed", type=int, default=42)
    args = parser.parse_args()

    random.seed(args.seed)
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "orders"
    ws.append(["order_id", "order_date", "customer", "sku", "qty", "amount", "cost"])

    start = date.today() - timedelta(days=30)
    for i in range(args.rows):
        qty = random.randint(1, 20)
        amount = round(qty * random.uniform(8, 30), 2)
        cost = round(amount * random.uniform(0.55, 0.9), 2)
        ws.append(
            [
                f"O{1000 + i}",
                (start + timedelta(days=random.randint(0, 29))).isoformat(),
                random.choice(CUSTOMERS),
                random.choice(SKUS),
                qty,
                amount,
                cost,
            ]
        )

    wb.save(out)
    print(json.dumps({"rows": args.rows, "file": str(out)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
