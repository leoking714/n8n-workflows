#!/usr/bin/env python3
"""读取 Excel，输出结构摘要（sheet、列名、行数、前 3 行）。"""

import argparse
import json
from pathlib import Path

from openpyxl import load_workbook


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--in", dest="src", required=True, help="输入 xlsx，如 /data/test/orders.xlsx")
    parser.add_argument("--out", dest="dst", required=True, help="摘要 JSON 输出路径")
    args = parser.parse_args()

    src = Path(args.src)
    if not src.exists():
        raise SystemExit(f"输入文件不存在：{src}")

    wb = load_workbook(src, read_only=True, data_only=True)
    summary = {"file": str(src), "sheets": {}}

    for name in wb.sheetnames:
        ws = wb[name]
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)

        count = 0
        sample = []
        for row in rows:
            if row and any(value is not None for value in row):
                count += 1
                if len(sample) < 3:
                    sample.append(list(row))

        summary["sheets"][name] = {
            "columns": [str(c) for c in (header or [])],
            "rows": count,
            "sample": sample,
        }

    wb.close()

    out = Path(args.dst)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    print(json.dumps({
        "ok": True,
        "file": str(out),
        "sheets": {k: v["rows"] for k, v in summary["sheets"].items()},
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
