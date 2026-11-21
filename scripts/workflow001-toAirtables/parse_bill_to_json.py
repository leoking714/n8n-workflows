#!/usr/bin/env python3
"""账单 Excel → JSON。用法：--in 必填，--out 可选，--sku 可选。"""

import argparse
import json
import sys
import warnings
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore", category=FutureWarning)


def _norm_order(v):
    """订单编号规范化。

    账单里的 Order Number 读出来可能是 int / float / str，
    计算表里可能是字符串，直接比较会因 "8027806499144" != 8027806499144 而匹配失败。
    这里统一转成去空白、去掉浮点尾巴 ".0" 的字符串；空值返回空串。
    """
    if pd.isna(v):
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def load_platform_skus(path: Path) -> pd.DataFrame:
    """读取计算表：订单编号 -> 平台SKU，按行顺序保留。

    返回两列 DataFrame：["Order Number", "Platform SKU"]。
    同一订单可能出现多行（主商品 + 配件），这里不去重、不聚合，
    保留原始行顺序，供后续按"第几行"对应到账单的"第几行"。

    注意：不指定 sheet 名，直接读工作簿默认打开的第一个 sheet。
    """
    df = pd.read_excel(path, sheet_name=0, header=0)

    order_col = "订单编号"
    sku_col = "平台SKU"
    if order_col not in df.columns or sku_col not in df.columns:
        raise ValueError(
            f"计算表缺少列：{order_col!r} 或 {sku_col!r}，"
            f"实际列：{list(df.columns)[:10]}"
        )

    sub = df[[order_col, sku_col]].copy()
    sub.columns = ["Order Number", "Platform SKU"]

    # 计算表 A 列里有 "=A12" 这类公式字符串（代表"同上一个订单"），
    # 先当成空，再用 ffill 把上一个真实订单号填下来。
    sub["Order Number"] = sub["Order Number"].apply(
        lambda v: pd.NA if isinstance(v, str) and v.strip().startswith("=") else v
    )
    sub["Order Number"] = sub["Order Number"].ffill()

    sub = sub.dropna(subset=["Order Number"])
    sub["Order Number"] = sub["Order Number"].map(_norm_order)
    return sub


def attach_platform_skus(df: pd.DataFrame, sku_df: pd.DataFrame) -> pd.DataFrame:
    """按 Order Number 顺序匹配平台SKU；同单号内后面的空行用前面最近值补齐。"""
    if "Order Number" not in df.columns:
        return df

    sku_df = sku_df.copy()
    sku_df["Order Number"] = sku_df["Order Number"].map(_norm_order)

    sku_map: dict[str, list] = {}
    for order, grp in sku_df.groupby("Order Number", sort=False):
        vals = [
            v for v in grp["Platform SKU"].tolist()
            if pd.notna(v) and str(v).strip() != ""
        ]
        sku_map[order] = vals

    out = df.copy()
    order_key = out["Order Number"].map(_norm_order)
    out["_order_idx"] = out.groupby(order_key, sort=False).cumcount()

    platform_skus = []
    for order, idx in zip(order_key, out["_order_idx"]):
        vals = sku_map.get(order, [])
        if not vals:
            platform_skus.append(pd.NA)
        elif idx < len(vals):
            platform_skus.append(vals[idx])
        else:
            platform_skus.append(vals[-1])

    out["Platform SKU"] = platform_skus
    out = out.drop(columns=["_order_idx"])

    out["Platform SKU"] = (
        out.groupby(order_key, sort=False)["Platform SKU"].ffill()
    )
    return out


def merge_by_transaction_sku(df: pd.DataFrame, store_col: str = "Store Name") -> pd.DataFrame:
    """按 (Order Number, Transaction Number, Platform SKU) 合并，只输出七个字段。

    关键修正：
      - 分组键加入 Order Number，避免不同订单共用同一交易号时被错误合并；
      - Store Name 按 Order Number 回填（店铺是订单级字段），不按交易号，
        防止同交易号多订单串店；
      - shipping fee 按 (Order Number, Transaction Number) 取第一个非空值，
        同订单内所有 SKU 共用该单运费；
      - Item Name 按分组取第一个非空值；
      - product cost 组内求和。
    """
    for c in ("Order Number", "Transaction Number", "Platform SKU"):
        if c not in df.columns:
            raise ValueError(f"合并需要列 {c!r}，实际列：{list(df.columns)[:15]}")

    cost_col = "Product cost"
    ship_col = "Shipping fee + Taxes"
    item_col = "Item Name"
    for c in (cost_col, ship_col, store_col, item_col):
        if c not in df.columns:
            raise ValueError(f"缺少列 {c!r}，实际列：{list(df.columns)[:15]}")

    work = df[
        [
            "Order Number",
            "Transaction Number",
            "Platform SKU",
            store_col,
            item_col,
            cost_col,
            ship_col,
        ]
    ].copy()

    work = work.dropna(subset=["Platform SKU"])
    work["_order_key"] = work["Order Number"].map(_norm_order)

    work["_cost"] = pd.to_numeric(work[cost_col], errors="coerce").fillna(0)
    work["_ship"] = pd.to_numeric(work[ship_col], errors="coerce")

    # 店铺：按订单取第一个非空值
    store_by_order = work.groupby("_order_key", sort=False)[store_col].first()

    # 运费：按 (订单, 交易号) 取第一个非空值，同一单内所有 SKU 共用
    ship_by_order_txn = (
        work.groupby(["_order_key", "Transaction Number"], sort=False)["_ship"].first()
    )

    # 分组聚合：product cost 求和，Item Name 取第一个非空
        # 先算 cost 求和
    cost_sum = (
        work.groupby(["_order_key", "Transaction Number", "Platform SKU"], sort=False)["_cost"]
        .sum()
        .rename("product cost")
    )
    # 再算 Item Name 第一个非空
    item_first = (
        work.dropna(subset=[item_col])
        .groupby(["_order_key", "Transaction Number", "Platform SKU"], sort=False)[item_col]
        .first()
        .rename("Item Name")
    )
    # 再算每组的原始行数（账单里同一 订单+交易号+SKU 出现了几行）
    # 账单里同订单同SKU多行会被合并成一行、金额合计；下游 Airtable 若同样有多条记录，
    # 就无法判断每条记录该填多少，需要人工处理。这里把行数带出去供流程判断/提醒。
    row_count = (
        work.groupby(["_order_key", "Transaction Number", "Platform SKU"], sort=False)
        .size()
        .rename("账单行数")
    )
    grouped = pd.concat([cost_sum, item_first, row_count], axis=1).reset_index()

    grouped["Store Name"] = grouped["_order_key"].map(store_by_order)
    grouped["shipping fee"] = grouped.set_index(
        ["_order_key", "Transaction Number"]
    ).index.map(ship_by_order_txn)

    grouped = grouped.rename(columns={"_order_key": "Order Number"})
    grouped = grouped[
        [
            "Order Number",
            "Transaction Number",
            "Store Name",
            "Item Name",
            "Platform SKU",
            "product cost",
            "shipping fee",
            "账单行数",
        ]
    ]

    return grouped.reset_index(drop=True)


def parse(path: Path, sku_path=None) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None)

    first = raw.iloc[:, 0].astype(str).str.strip()
    hits = first[first == "No."]
    if hits.empty:
        raise ValueError("找不到表头行（第一列为 'No.' 的行）")

    df = pd.read_excel(path, header=int(hits.index[0]))

    if "Item Name" not in df.columns:
        raise ValueError(f"缺少 'Item Name' 列，实际列：{list(df.columns)[:10]}")

    mask = df["Item Name"].isna()
    if mask.any():
        df = df.iloc[: mask.idxmax()]

    # 注意：ffill 会把配件行空着的 Transaction Number 填成上一单值。
    # 这里保留原行为，但合并阶段已改为按 Order Number 分组，避免串单。
    _cost_backup = df["Product cost"].copy()
    df = df.ffill()
    df["Product cost"] = _cost_backup

    for col in ("Transaction Number"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="raise").astype("Int64")

    store_col = None
    for candidate in ("Store Name", "Store"):
        if candidate in df.columns:
            store_col = candidate
            break
    if store_col is None:
        raise ValueError(
            f"缺少门店列（'Store Name' 或 'Store'），实际列：{list(df.columns)[:15]}"
        )

    if sku_path is not None:
        sku_df = load_platform_skus(Path(sku_path))
        df = attach_platform_skus(df, sku_df)

    df = merge_by_transaction_sku(df, store_col=store_col)
    return df



def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--in", dest="src", required=True)
    parser.add_argument("--out", dest="dst", help="输出 JSON 文件；不传则打印到 stdout")
    parser.add_argument("--sku", dest="sku_src", help="计算表 Excel（含订单编号/平台SKU）")
    args = parser.parse_args()

    src = Path(args.src)
    try:
        if not src.exists():
            raise FileNotFoundError(f"文件不存在：{src}")

        sku_path = Path(args.sku_src) if args.sku_src else None
        if sku_path is not None and not sku_path.exists():
            raise FileNotFoundError(f"SKU 文件不存在：{sku_path}")

        df = parse(src, sku_path)
        records = json.loads(df.to_json(orient="records", force_ascii=False))

        if args.dst:
            dst = Path(args.dst)
            dst.parent.mkdir(parents=True, exist_ok=True)
            dst.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
            result = {"ok": True, "rows": len(records), "file": str(dst)}
        else:
            result = {"ok": True, "rows": len(records), "data": records}

        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0

    except Exception as exc:
        print(json.dumps({
            "ok": False,
            "file": str(src),
            "error": f"{type(exc).__name__}: {exc}",
        }, ensure_ascii=False, indent=2))

        return 1


if __name__ == "__main__":
    sys.exit(main())