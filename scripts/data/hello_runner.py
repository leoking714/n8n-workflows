#!/usr/bin/env python3
"""最小示例：接收参数，生成一个 JSON 文件，并返回摘要。"""

import argparse
import json
from datetime import datetime
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--name", default="world", help="要打招呼的名字")
    parser.add_argument("--out", required=True, help="输出文件，如 /data/example/hello.json")
    args = parser.parse_args()

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)

    payload = {
        "hello": args.name,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "cwd": str(Path.cwd()),
    }
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    # stdout 只输出结果 JSON
    print(json.dumps({"ok": True, "name": args.name, "file": str(out)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
