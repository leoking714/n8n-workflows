#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""读取 Excel，输出每行第 1 列和第 5 列的 JSON。

用法：
    python first_fifth_to_json.py --in orders.xlsx
    python first_fifth_to_json.py --in orders.xlsx --skip-header
    python first_fifth_to_json.py --in orders.xlsx --sheet 0 --out result.json

输出：
    [
      {"column1": "O10000", "column5": 2},
      ...
    ]

说明：
- 默认包含所有行；有表头时加 --skip-header。
- 两列都为空的尾随空行会跳过。
- 只支持 .xlsx/.xlsm；.xls 请先另存为 .xlsx。
"""

import argparse
import json
import sys
from datetime import date, datetime, time
from pathlib import Path

from openpyxl import load_workbook


def json_value(value):
    """把 Excel 值转成可 JSON 序列化的值。"""
    if value is None:
        return None
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    return value


def get_sheet(wb, sheet_arg):
    """按名称或序号选择 worksheet；默认活动表。"""
    if sheet_arg is None:
        return wb.active
    if sheet_arg.isdigit():
        idx = int(sheet_arg)
        if idx < 0 or idx >= len(wb.sheetnames):
            raise SystemExit(f"sheet 序号越界: {sheet_arg}")
        return wb[wb.sheetnames[idx]]
    if sheet_arg not in wb.sheetnames:
        raise SystemExit(f"找不到 sheet: {sheet_arg}")
    return wb[sheet_arg]


def main():
    parser = argparse.ArgumentParser(description="Excel 第1/5列转 JSON")
    parser.add_argument("--in", dest="input_path", required=True, help="Excel 文件路径")
    parser.add_argument("--sheet", default=None, help="sheet 名称或序号；默认活动表")
    parser.add_argument("--out", default=None, help="输出 JSON 路径；不传则输出到 stdout")
    parser.add_argument("--skip-header", action="store_true", help="跳过第一行表头")
    args = parser.parse_args()

    path = Path(args.input_path)
    if not path.exists():
        raise SystemExit(f"文件不存在: {path}")
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise SystemExit("只支持 .xlsx/.xlsm；.xls 请先另存为 .xlsx")

    wb = load_workbook(path, read_only=True, data_only=True)
    ws = get_sheet(wb, args.sheet)

    rows = []
    for index, row in enumerate(ws.iter_rows(min_col=1, max_col=5, values_only=True), start=1):
        if args.skip_header and index == 1:
            continue

        first = row[0] if len(row) >= 1 else None
        fifth = row[4] if len(row) >= 5 else None

        # 两列都为空的行不输出，避免 Excel 尾部空白行
        if first in (None, "") and fifth in (None, ""):
            continue

        rows.append({
            "column1": json_value(first),
            "column5": json_value(fifth),
        })

    wb.close()

    text = json.dumps(rows, ensure_ascii=False, indent=2)
    if args.out:
        Path(args.out).write_text(text, encoding="utf-8")
        print(f"written {len(rows)} rows -> {args.out}", file=sys.stderr)
    else:
        print(text)

    return 0


if __name__ == "__main__":
    sys.exit(main())
