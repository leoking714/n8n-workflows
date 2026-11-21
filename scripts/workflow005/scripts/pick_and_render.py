#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
核对后生成 Bill：
  1) 定位计算表（指定文件名，或 output 目录里最新）
  2) 生成 PDF（docx 模板 / 程序绘制两种引擎）

用法：
  python pick_and_render.py '{"file":"", "date":"2026. 11. 20", "dir":"/data/...", "engine":"docx"}'
"""
import glob, json, os, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))


def pick(cfg):
    d = cfg.get("dir") or "/data/data/workflow005/output"
    name = (cfg.get("file") or "").strip()
    if name:
        p = name if os.path.isabs(name) else os.path.join(d, name)
        if not os.path.exists(p):
            raise SystemExit(f"找不到文件：{p}")
        return p
    cands = sorted(glob.glob(os.path.join(d, "60103计算表_*.xlsx")), key=os.path.getmtime)
    if not cands:
        raise SystemExit("output 目录里没有 60103计算表_*.xlsx")
    return cands[-1]


def main():
    cfg = json.loads(sys.argv[1]) if len(sys.argv) > 1 else {}
    xlsx = pick(cfg)
    engine = (cfg.get("engine") or "docx").lower()
    out = os.path.splitext(xlsx)[0] + ("_bill.pdf" if engine == "docx" else ".pdf")

    script = "bill_docx.py" if engine == "docx" else "bill_pdf_from_xlsx.py"
    cmd = [sys.executable, os.path.join(HERE, script), "--xlsx", xlsx, "--out", out]
    if cfg.get("date"):
        cmd += ["--date", cfg["date"]]
    lb = cfg.get("last_balance")
    if lb not in (None, ""):
        cmd += ["--last-balance", str(lb)]
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=900)
    if r.returncode != 0:
        print(r.stderr[-1500:], file=sys.stderr)
        raise SystemExit(1)
    line = [l for l in r.stdout.strip().split("\n") if l.strip().startswith("{")][-1]
    res = json.loads(line)
    res["xlsx"] = xlsx
    res["engine"] = engine
    print(json.dumps(res, ensure_ascii=False))


if __name__ == "__main__":
    main()
