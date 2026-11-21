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
    # 默认读取第一个 sheet（即 Excel 打开时默认显示的那个）
    df = pd.read_excel(path, sheet_name=0, header=0)

    order_col = "订单编号"
    sku_col = "平台SKU"
    if order_col not in df.columns or sku_col not in df.columns:
        raise ValueError(
            f"计算表缺少列：{order_col!r} 或 {sku_col!r}，"
            f"实际列：{list(df.columns)[:10]}"
        )

    # 只取需要的两列，并重命名成与账单一致的列名
    sub = df[[order_col, sku_col]].copy()
    sub.columns = ["Order Number", "Platform SKU"]

    # 计算表 A 列里有 "=A12" 这类公式字符串（代表"同上一个订单"），
    # 先当成空，再用 ffill 把上一个真实订单号填下来。
    sub["Order Number"] = sub["Order Number"].apply(
        lambda v: pd.NA if isinstance(v, str) and v.strip().startswith("=") else v
    )
    sub["Order Number"] = sub["Order Number"].ffill()

    # 去掉仍然没有订单号的行（表尾空行等）
    sub = sub.dropna(subset=["Order Number"])
    sub["Order Number"] = sub["Order Number"].map(_norm_order)
    return sub


def attach_platform_skus(df: pd.DataFrame, sku_df: pd.DataFrame) -> pd.DataFrame:
    """按 Order Number 顺序匹配平台SKU；同单号内后面的空行用前面最近值补齐。

    规则：
      1. 同一订单在计算表里按出现顺序有多个 SKU；
      2. 账单里同一订单的第 n 行，取计算表该订单的第 n 个 SKU；
      3. 账单行数超过 SKU 数时，多出的行先用最后一个 SKU；
      4. 最后再按订单分组 ffill，保证"后面没有 SKU 就补前面最近一个"。
    """
    if "Order Number" not in df.columns:
        return df

    sku_df = sku_df.copy()
    sku_df["Order Number"] = sku_df["Order Number"].map(_norm_order)

    # 同一订单按出现顺序收集非空 SKU，形成 {订单号: [SKU1, SKU2, ...]}
    sku_map: dict[str, list] = {}
    for order, grp in sku_df.groupby("Order Number", sort=False):
        vals = [
            v for v in grp["Platform SKU"].tolist()
            if pd.notna(v) and str(v).strip() != ""
        ]
        sku_map[order] = vals

    out = df.copy()

    # 把账单的 Order Number 规范化后作为分组键
    order_key = out["Order Number"].map(_norm_order)

    # 计算每个订单在账单中的行内序号（第 0 行、第 1 行……），
    # 用来决定取计算表里的第几个 SKU。
    out["_order_idx"] = out.groupby(order_key, sort=False).cumcount()

    # 逐行分配 SKU：序号在范围内取对应 SKU，超出取最后一个，没有则 NA
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

    # 同单号内：空值用前面最近的非空值补齐
    # （处理账单里配件行在计算表中没有对应 SKU 的情况）
    out["Platform SKU"] = (
        out.groupby(order_key, sort=False)["Platform SKU"].ffill()
    )
    return out


def merge_by_transaction_sku(df: pd.DataFrame) -> pd.DataFrame:
    """按 (Transaction Number, Platform SKU) 合并，只输出五个字段。

    规则：
      - product cost：组内 Total product cost 求和；
      - shipping fee：同一 Transaction Number 内所有行共用该交易号的运费总计，
        即组内取第一个非空值（账单里同一交易号各行运费相同）；
      - Store Name：订单级字段，同一交易号内取第一个非空值；
      - 输出列：Transaction Number / Store Name / Platform SKU /
                product cost / shipping fee。
    """
    if "Transaction Number" not in df.columns or "Platform SKU" not in df.columns:
        raise ValueError(
            "合并需要 'Transaction Number' 和 'Platform SKU' 列，"
            f"实际列：{list(df.columns)[:15]}"
        )

    cost_col = "Product cost"
    ship_col = "Shipping fee + Taxes"
    store_col = "Store Name"
    for c in (cost_col, ship_col, store_col):
        if c not in df.columns:
            raise ValueError(f"缺少列 {c!r}，实际列：{list(df.columns)[:15]}")

    work = df[
        ["Transaction Number", "Platform SKU", store_col, cost_col, ship_col]
    ].copy()

    # 空 SKU 不参与合并，避免把不同商品误并到一起
    work = work.dropna(subset=["Platform SKU"])

    # 数值化，空值当 0 便于求和；运费保留 NaN 以便取第一个非空
    work["_cost"] = pd.to_numeric(work[cost_col], errors="coerce").fillna(0)
    work["_ship"] = pd.to_numeric(work[ship_col], errors="coerce")

    # 订单级字段：同一交易号取第一个非空值
    ship_by_txn = (
        work.groupby("Transaction Number", sort=False)["_ship"]
        .first()
    )
    store_by_txn = (
        work.groupby("Transaction Number", sort=False)[store_col]
        .first()
    )

    # 按 (Transaction Number, Platform SKU) 聚合 product cost
    grouped = (
        work.groupby(["Transaction Number", "Platform SKU"], sort=False)["_cost"]
        .sum()
        .reset_index()
    )

    # 回填订单级字段
    grouped["Store Name"] = grouped["Transaction Number"].map(store_by_txn)
    grouped["shipping fee"] = grouped["Transaction Number"].map(ship_by_txn)

    # 重命名并只保留五个字段
    grouped = grouped.rename(columns={"_cost": "product cost"})
    grouped = grouped[
        ["Transaction Number", "Store Name", "Platform SKU", "product cost", "shipping fee"]
    ]

    return grouped.reset_index(drop=True)


def parse(path: Path, sku_path=None) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None)

    # 先不设表头，扫描第一列找 "No." 所在行，作为真正的表头行
    first = raw.iloc[:, 0].astype(str).str.strip()
    hits = first[first == "No."]
    if hits.empty:
        raise ValueError("找不到表头行（第一列为 'No.' 的行）")

    df = pd.read_excel(path, header=int(hits.index[0]))

    if "Item Name" not in df.columns:
        raise ValueError(f"缺少 'Item Name' 列，实际列：{list(df.columns)[:10]}")

    # 商品行到第一个 Item Name 为空的行为止（后面是 Payment fee / Total 等合计区）
    mask = df["Item Name"].isna()
    if mask.any():
        df = df.iloc[: mask.idxmax()]

    df = df.ffill()

    # for col in ("Order Number", "Transaction Number"):
    for col in ("Transaction Number"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="raise").astype("Int64")

    # 新增：合并平台SKU（不传 --sku 时保持原行为）
    if sku_path is not None:
        sku_df = load_platform_skus(Path(sku_path))
        df = attach_platform_skus(df, sku_df)

    # 新增：按 (Transaction Number, Platform SKU) 合并并只输出四个字段
    df = merge_by_transaction_sku(df)

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

        # 可选 SKU 文件：传了才做合并，传了但不存在则报错
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