# -*- coding: utf-8 -*-
"""
校验业务填好的「商品配置-待填模板」，通过后可直接导入飞书多维表格。

用法：
    python 校验商品配置模板.py 填好的.xlsx
    python 校验商品配置模板.py 填好的.xlsx --json 输出/商品配置.json
"""
import argparse
import json
import os
import sys

import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
PROJ = os.path.dirname(BASE)
sys.path.insert(0, PROJ)

NUM_COLS = ["采购价", "采购价_美国偏远", "国内运费", "国内运费_美国偏远", "重量",
            "服务费_美国", "服务费", "运价覆盖_公斤价", "运价覆盖_挂号费",
            "账单固定金额", "每计费单位件数"]
REQUIRED = ["采购价", "国内运费", "重量", "服务费", "每计费单位件数", "状态"]

HEAD_MAP = {
    "SKU前缀": "SKU前缀",
    "采购价_默认": "采购价",
    "采购价_美国偏远": "采购价_美国偏远",
    "国内运费_默认": "国内运费",
    "国内运费_美国偏远": "国内运费_美国偏远",
    "重量(KG)": "重量",
    "服务费_美国": "服务费_美国",
    "服务费_其他": "服务费",
    "运价覆盖_国家": "运价覆盖_国家",
    "运价覆盖_公斤价": "运价覆盖_公斤价",
    "运价覆盖_挂号费": "运价覆盖_挂号费",
    "账单固定金额$": "账单固定金额",
    "每计费单位件数": "每计费单位件数",
    "状态": "状态",
    "填写人": "填写人",
    "备注": "备注",
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--json", help="校验通过后导出 JSON（可直接推到飞书/给 calc 用）")
    args = ap.parse_args()

    df = pd.read_excel(args.xlsx, sheet_name="商品配置", dtype=object)
    df.columns = [str(c).strip() for c in df.columns]

    problems = []

    # ① 表头
    miss = [c for c in HEAD_MAP if c not in df.columns]
    if miss:
        problems.append(f"表头缺列：{miss}")
    df = df.rename(columns={k: v for k, v in HEAD_MAP.items() if k in df.columns})

    if "SKU前缀" not in df.columns:
        print("❌ 找不到 SKU前缀 列，无法继续")
        return 1

    df["SKU前缀"] = df["SKU前缀"].astype(str).str.strip()
    df = df[df["SKU前缀"].ne("") & df["SKU前缀"].ne("nan")].copy()

    # ② 数字列
    for c in NUM_COLS:
        if c not in df.columns:
            continue
        raw = df[c]
        df[c] = pd.to_numeric(raw, errors="coerce")
        bad = raw.notna() & df[c].isna()
        for i in df.index[bad]:
            problems.append(f"第 {i + 2} 行「{c}」不是数字：{raw.loc[i]!r}")

    # ③ 必填
    for c in REQUIRED:
        if c not in df.columns:
            problems.append(f"缺少必填列「{c}」")
            continue
        if c == "状态":
            bad = df[c].isna() | ~df[c].astype(str).str.strip().isin(["已确认", "待确认"])
        else:
            bad = df[c].isna()
        for i in df.index[bad]:
            problems.append(f"第 {i + 2} 行「{c}」未填")

    # ④ 业务规则
    for i, r in df.iterrows():
        row = i + 2
        if pd.notna(r.get("重量")):
            wt = r["重量"]
            is_ng = str(r.get("SKU前缀", "")).upper().startswith("SHIP")
            lo_ok = (wt >= 0) if is_ng else (wt > 0)      # 非实物行允许 0
            if not (lo_ok and wt <= 20):
                problems.append(f"第 {row} 行重量 {wt} 不合理（0~20，非实物 SKU 可为 0）")
        if pd.notna(r.get("服务费")) and not (1 <= r["服务费"] <= 2):
            problems.append(f"第 {row} 行服务费 {r['服务费']} 不在 1~2")
        p = r.get("每计费单位件数")
        if pd.notna(p) and (p < 1 or p != int(p)):
            problems.append(f"第 {row} 行每计费单位件数 {p} 必须是 ≥1 的整数")
        if pd.notna(r.get("运价覆盖_国家")) and (pd.isna(r.get("运价覆盖_公斤价"))):
            problems.append(f"第 {row} 行填了运价覆盖国家，但没填公斤价")

    # ⑤ 重复
    # 一格里可以写多个前缀（英文/中文逗号、顿号），展开后查重
    import re as _re
    alias_seen, alias_dup = {}, []
    for i, cell in df["SKU前缀"].items():
        for a in _re.split(r"[,，、]", str(cell)):
            a = a.strip()
            if not a:
                continue
            if a in alias_seen:
                alias_dup.append((a, alias_seen[a], i + 2))
            else:
                alias_seen[a] = i + 2
    if alias_dup:
        problems.append("前缀别名重复：" +
                        "；".join(f"{a}（第{r1}行和第{r2}行）" for a, r1, r2 in alias_dup))

    # ⑥ 前缀格式
    for i, v in df["SKU前缀"].items():
        for a in _re.split(r"[,，、]", str(v)):
            a = a.strip()
            if not a:
                continue
            if " " in a:
                problems.append(f"第 {i + 2} 行「{a}」含空格")
            if a.islower():
                problems.append(f"第 {i + 2} 行「{a}」建议大写（和小写不是同一个）")

    if problems:
        print(f"❌ 发现 {len(problems)} 个问题：")
        for p in problems[:50]:
            print("   -", p)
        if len(problems) > 50:
            print(f"   …… 还有 {len(problems) - 50} 条")
        return 1

    print(f"✅ 校验通过：{len(df)} 个 SKU 前缀")

    if args.json:
        out = []
        for _, r in df.iterrows():
            item = {"SKU前缀": r["SKU前缀"]}
            for c in NUM_COLS:
                if c in df.columns:
                    v = r.get(c)
                    item[c] = None if pd.isna(v) else float(v)
                else:
                    item[c] = None
            for c in ("运价覆盖_国家", "状态", "填写人", "备注"):
                v = r.get(c)
                item[c] = None if pd.isna(v) else str(v).strip()
            out.append(item)
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        print(f"已导出 {args.json}")

    stat = df.get("状态", pd.Series(dtype=str))
    if len(stat):
        print("状态分布：", stat.value_counts().to_dict())
    return 0


if __name__ == "__main__":
    sys.exit(main())
