# -*- coding: utf-8 -*-
"""
把 n8n 解析好的订单明细写成带格式的 Excel（供人工查看/对账）。

用法：
    python write_orders_xlsx.py --input /data/data/workflow005/input/orders_run.json \
                                --out   /data/data/workflow005/output/订单导出_xxx.xlsx
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from export_orders import write_xlsx  # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--start", default="")
    ap.add_argument("--end", default="")
    args = ap.parse_args()

    with open(args.input, encoding="utf-8") as f:
        payload = json.load(f)

    rows = payload.get("rows") if isinstance(payload, dict) else payload
    if not rows:
        print("RESULT " + json.dumps({"ok": False, "error": "订单为空"}, ensure_ascii=False))
        return 1

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    write_xlsx(rows, args.out,
               args.start or (payload.get("start") if isinstance(payload, dict) else "") or "",
               args.end or (payload.get("end") if isinstance(payload, dict) else "") or "")

    orders = len({str(r.get("交易编号")) for r in rows})
    print("RESULT " + json.dumps({
        "ok": True, "out": os.path.abspath(args.out), "rows": len(rows), "orders": orders,
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
