# -*- coding: utf-8 -*-
"""
一步步做匹配，把「匹配不上」的东西单独吐出来，交给业务补填。

步骤（按顺序，前一步不通就不往下走）：
    ① 国家改写        美国+岛屿→美国（偏远）；澳大利亚→分区
    ② SKU 前缀匹配     没匹配到 → 待补商品配置
    ③ 运价段匹配       国家 + 整单重量 找不到段 → 待补运价
    ④ 只做统计，不算钱

输出（stdout 最后一行）：
    RESULT {"ok":true, "steps":[...], "missing_sku":[...], "missing_rate":[...]}

用法：
    python step_match.py
    python step_match.py --order /data/.../订单导出数据_xxx.xlsx
    python step_match.py --data-dir /data/data/workflow003/60103飞书/data
"""
import argparse
import glob
import json
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

import calc   # noqa: E402
import paths  # noqa: E402

DEFAULT_INPUT = next((p for p in ("/data/data/workflow003/input",
                                  "/Users/zuo/n8n/files/data/workflow003/input")
                      if os.path.isdir(p)), "/data/data/workflow003/input")


def newest_order(folder):
    files = sorted(glob.glob(os.path.join(folder, "订单导出数据_*.xlsx")),
                   key=os.path.getmtime)
    return files[-1] if files else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--order")
    ap.add_argument("--orders-json", help="ERP 拉下来的订单 JSON（优先级高于 --order）")
    ap.add_argument("--input-dir", default=DEFAULT_INPUT)
    ap.add_argument("--data-dir")
    args = ap.parse_args()

    if args.orders_json:
        order = args.orders_json
        if not os.path.exists(order):
            print("RESULT " + json.dumps({"ok": False, "error": f"订单 JSON 不存在：{order}"},
                                         ensure_ascii=False))
            return 1
        cfg = calc.load_cfg(args.data_dir)
        df = calc.read_orders_json(order)
    else:
        order = args.order or newest_order(args.input_dir)
        if not order or not os.path.exists(order):
            print("RESULT " + json.dumps({"ok": False, "error": "找不到订单文件"},
                                         ensure_ascii=False))
            return 1
        cfg = calc.load_cfg(args.data_dir)
        df = calc.read_orders(order)

    steps = []

    # ---------- ① 国家改写 ----------
    df = calc.fix_country(df, cfg)
    au_bad = int((df["国家"] == "澳大利亚").sum())
    steps.append({
        "step": 1, "name": "国家改写",
        "detail": f"订单 {df['交易编号'].nunique()} 单 / {len(df)} 行"
                  + (f"；澳洲邮编查不到分区 {au_bad} 行" if au_bad else ""),
        "warn": au_bad,
    })

    # ---------- ② SKU 前缀匹配 ----------
    prefixes = sorted(cfg["products"], key=len, reverse=True)

    def guess_prefix(sku):
        head = str(sku).split("_")[0]
        parts = head.split("-")
        return "-".join(parts[:2]) if len(parts) >= 2 else head

    def find_prefix(sku):
        s = str(sku).strip()
        for p in prefixes:
            if s == p or s.startswith(p + "-") or s.startswith(p + "_"):
                return p
        return None

    df["前缀"] = df["SKU"].map(find_prefix)
    unmatched = df[df["前缀"].isna()]

    missing_sku = []
    if len(unmatched):
        unmatched = unmatched.copy()
        unmatched["建议前缀"] = unmatched["SKU"].map(guess_prefix)
        agg = (unmatched.groupby("建议前缀")
               .agg(订单数=("交易编号", "nunique"), 行数=("SKU", "size"),
                    SKU数=("SKU", "nunique"))
               .sort_values("订单数", ascending=False))
        for pfx, r in agg.iterrows():
            skus = sorted(unmatched.loc[unmatched["建议前缀"] == pfx, "SKU"].unique())
            sample = unmatched.loc[unmatched["建议前缀"] == pfx, "交易编号"].iloc[0]
            missing_sku.append({
                "类型": "SKU",
                "标识": pfx,
                "建议前缀": pfx,
                "订单数": int(r["订单数"]),
                "行数": int(r["行数"]),
                "SKU数": int(r["SKU数"]),
                "样例交易编号": str(sample),
                "SKU样例": "、".join(skus[:3]) + ("…" if len(skus) > 3 else ""),
                "说明": f"商品配置里没有前缀「{pfx}」，需补一行（{int(r['SKU数'])} 个尺码在用）",
            })

    steps.append({
        "step": 2, "name": "SKU 前缀匹配",
        "detail": f"匹配上 {int(df['前缀'].notna().sum())} 行；未匹配 "
                  f"{len(unmatched)} 行 / {len(missing_sku)} 个 SKU",
        "warn": len(missing_sku),
    })

    # 没匹配上的 SKU 用兜底值继续算重量（只为能查运价，不影响补填清单）
    df["采购价"] = df["前缀"].map(lambda p: (cfg["products"].get(p) or {}).get("采购价"))
    df["重量"] = df["前缀"].map(lambda p: (cfg["products"].get(p) or {}).get("重量"))
    df["重量"] = df["重量"].fillna(calc.P(cfg, "WEIGHT_FALLBACK", 0.17))
    df["行重量"] = df["商品数量"] * df["重量"]
    df["订单重量"] = (df.groupby("交易编号", sort=False)["行重量"]
                      .transform("sum").round(3))

    # ---------- ③ 运价段匹配 ----------
    heads = (df.loc[df["首行"], ["交易编号", "国家", "订单重量"]]
               .drop_duplicates("交易编号"))
    pairs = heads.apply(lambda r: calc.price_of(cfg, r["国家"], r["订单重量"]), axis=1)
    heads["公斤价"] = [x[0] for x in pairs]
    heads["挂号费"] = [x[1] for x in pairs]
    no_rate = heads[heads["公斤价"].isna()]

    missing_rate = []
    for _, r in no_rate.iterrows():
        ctry, w = r["国家"], float(r["订单重量"])
        segs = cfg["rates"].get(ctry)
        if not segs:
            need = f"该国家在运价表里没有任何段，需要新增「{ctry}」的运价"
            segs_txt = "无"
        else:
            lo, hi = segs[0][0], segs[-1][1]
            need = f"需要覆盖 {w}kg 的段（现有 {lo}~{hi}kg）"
            segs_txt = "、".join(f"{a}-{b}" for a, b, *_ in segs)
        missing_rate.append({
            "类型": "运价",
            "标识": f"{ctry}|{w}kg",
            "建议前缀": None,
            "订单数": int((df["交易编号"] == r["交易编号"]).sum()),
            "行数": 1,
            "样例交易编号": str(r["交易编号"]),
            "说明": need,
        })

    steps.append({
        "step": 3, "name": "运价匹配",
        "detail": f"订单 {len(heads)} 单；缺运价 {len(no_rate)} 单",
        "warn": len(no_rate),
    })

    # ---------- ④ 汇总 ----------
    result = {
        "ok": True,
        "order": os.path.basename(order),
        "source": "erp_json" if args.orders_json else "xlsx",
        "orders": int(df["交易编号"].nunique()),
        "rows": int(len(df)),
        "missing_count": len(missing_sku) + len(missing_rate),
        "steps": steps,
        "missing_sku": missing_sku,
        "missing_rate": missing_rate,
    }
    print(f"订单：{result['order']}")
    for s in steps:
        print(f"  步骤{s['step']} {s['name']}: {s['detail']}")
    print(f"待补：SKU {len(missing_sku)} 个 / 运价 {len(missing_rate)} 单")
    print("RESULT " + json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
