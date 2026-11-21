# -*- coding: utf-8 -*-
"""
60103 生成计算表（新工程 · 飞书版）

流程：
    订单导出 xlsx ─┐
                  ├→ calc.py（查配置、算钱、出异常）→ writer.py（写 xlsx）
    data/*.json  ─┘

用法：
    python generate.py                              # 用默认订单（最新导出）
    python generate.py 订单.xlsx
    python generate.py 订单.xlsx --out 输出/xx.xlsx
    python generate.py --data-dir data
"""
import argparse
import glob
import os
import sys
from pathlib import Path

import calc
import paths
import writer

BASE = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.environ.get("BILL_INPUT") or next(
    (p for p in ("/data/data/workflow003/input",
                 "/Users/zuo/n8n/files/data/workflow003/input") if os.path.isdir(p)),
    "/data/data/workflow003/input")
OUT_DIR = paths.out_dir()


def newest_order(folder=INPUT_DIR):
    files = sorted(glob.glob(os.path.join(folder, "订单导出数据_*.xlsx")),
                   key=os.path.getmtime)
    return files[-1] if files else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("order", nargs="?", help="订单导出 xlsx；不填则取最新")
    ap.add_argument("--orders-json", help="ERP 拉下来的订单 JSON（优先）")
    ap.add_argument("--out", help="输出路径")
    ap.add_argument("--data-dir", help="配置 JSON 目录（默认 data/）")
    args = ap.parse_args()

    order = args.orders_json or args.order or newest_order()
    if not order or not os.path.exists(order):
        print(f"找不到订单文件：{order}")
        return 1

    print(f"订单：{order}（{'ERP JSON' if args.orders_json else 'xlsx'}）")
    if args.orders_json:
        cfg = calc.load_cfg(args.data_dir)
        df = calc.read_orders_json(order)
        df = calc.fix_country(df, cfg)
        df = calc.match_sku(df, cfg)
        df = calc.match_rate(df, cfg)
        df = calc.money(df, cfg)
        issues = calc.build_issues(df, cfg)
        agg, total = calc.summarize(df, cfg)
    else:
        df, issues, agg, total = calc.run(order, args.data_dir)

    os.makedirs(OUT_DIR, exist_ok=True)
    out = args.out or os.path.join(
        OUT_DIR, f"60103计算表_{Path(order).stem.replace('订单导出数据_', '')}.xlsx")

    writer.write_excel(df, issues, out)

    print(f"明细 {len(df)} 行 / {df['交易编号'].nunique()} 单；异常 {len(issues)} 条")
    print(f"汇总 {total[1]} 单 / {total[2]} 件 / 账单 {total[-1]}$")
    print(f"已生成 {out}")

    # 给 n8n 解析的最后一行
    import json
    print("RESULT " + json.dumps({
        "ok": True,
        "orders": int(df["交易编号"].nunique()),
        "rows": int(len(df)),
        "issues": int(len(issues)),
        "bill_total": float(total[-1]),
        "out": os.path.abspath(out),
        "source": "erp_json" if args.orders_json else "xlsx",
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
