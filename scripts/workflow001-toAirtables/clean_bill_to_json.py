#!/usr/bin/env python3
"""账单 Excel → JSON。用法：--in 必填，--out 可选。"""

import argparse
import json
import sys
import warnings
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore", category=FutureWarning)


def parse(path: Path) -> pd.DataFrame:
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

    df = df.ffill()

    # for col in ("Order Number", "Transaction Number"):
    for col in ("Transaction Number"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="raise").astype("Int64")

    return df


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--in", dest="src", required=True)
    parser.add_argument("--out", dest="dst", help="输出 JSON 文件；不传则打印到 stdout")
    args = parser.parse_args()

    src = Path(args.src)
    try:
        if not src.exists():
            raise FileNotFoundError(f"文件不存在：{src}")

        df = parse(src)
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