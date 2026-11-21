# -*- coding: utf-8 -*-
"""
把飞书多维表格的五张配置表拉到本地，存成 JSON。

新工程的「数据同步」环节：只负责取数 + 落盘 + 体检，不做金额计算。

用法：
    export FEISHU_APP_ID=cli_xxx
    export FEISHU_APP_SECRET=xxx
    export FEISHU_APP_TOKEN=bascnxxx
    python sync_config.py                 # 存到 data/
    python sync_config.py --check         # 只体检，不落盘
"""
import argparse
import json
import os
import sys

import paths
from feishu_client import Feishu, fetch_all

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data")

# 每张表必须有的字段（缺了直接报错，避免跑到一半才发现）
REQUIRED = {
    "商品配置": ["SKU前缀", "采购价", "国内运费", "重量"],
    "运价表": ["国家", "下界", "上界", "公斤价", "挂号费"],
    "澳洲分区": ["邮编", "分区"],
    "全局参数": ["参数键", "参数值"],
    "国家规则": ["国家", "是否欧盟"],
}


def check(data):
    """体检：字段齐不齐、主键重不重复、值合不合理"""
    problems = []

    for table, cols in REQUIRED.items():
        rows = data.get(table, [])
        if not rows:
            problems.append(f"{table}: 一条记录都没有（表名对不对？应用加协作者了吗？）")
            continue
        keys = set(rows[0].keys())
        miss = [c for c in cols if c not in keys]
        if miss:
            problems.append(f"{table}: 缺少字段 {miss}；实际字段 = {sorted(keys)}")

    # 商品配置：SKU 前缀唯一
    p = data.get("商品配置", [])
    seen, dup = set(), []
    for r in p:
        k = str(r.get("SKU前缀") or "").strip()
        if not k:
            problems.append("商品配置: 有空 SKU前缀")
            continue
        if k in seen:
            dup.append(k)
        seen.add(k)
    if dup:
        problems.append(f"商品配置: SKU前缀 重复 {sorted(set(dup))}")

    # 运价表：下界/上界 必须是数字
    bad = [r for r in data.get("运价表", [])
           if not isinstance(r.get("下界"), (int, float))
           or not isinstance(r.get("上界"), (int, float))]
    if bad:
        problems.append(f"运价表: {len(bad)} 行下界/上界不是数字（字段类型要设成「数字」）")

    # 澳洲分区：邮编必须是 4 位文本
    bad_au = [r for r in data.get("澳洲分区", [])
              if not (isinstance(r.get("邮编"), str) and len(r["邮编"]) == 4)]
    if bad_au:
        problems.append(f"澳洲分区: {len(bad_au)} 行邮编不是 4 位文本"
                        f"（字段类型要设成「文本」，否则前导零会丢）")

    return problems


def summary(data):
    lines = []
    for name, rows in data.items():
        lines.append(f"  {name:8s} {len(rows):>5} 行")
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="只体检，不写文件")
    ap.add_argument("--app-id", help="飞书 App ID（不填读环境变量）")
    ap.add_argument("--app-secret", help="飞书 App Secret")
    ap.add_argument("--app-token", help="多维表格 app_token")
    ap.add_argument("--out", help="输出目录（默认自动选可写位置）")
    args = ap.parse_args()

    data = fetch_all(Feishu(app_id=args.app_id, app_secret=args.app_secret,
                            app_token=args.app_token))
    print("拉到：")
    print(summary(data))

    problems = check(data)
    if problems:
        print("\n⚠️ 体检发现问题：")
        for p in problems:
            print("   -", p)
        if args.check:
            return 1
    else:
        print("\n✅ 体检通过")

    if args.check:
        return 0

    args.out = args.out or paths.data_dir()
    os.makedirs(args.out, exist_ok=True)
    for name, rows in data.items():
        path = os.path.join(args.out, f"{name}.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(rows, f, ensure_ascii=False, indent=2)
    print(f"\n已写入 {args.out}/")
    return 0


if __name__ == "__main__":
    sys.exit(main())
