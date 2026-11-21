# -*- coding: utf-8 -*-
"""
生成「商品配置-待填模板」给业务填写。

做法：
    1. 扫订单导出（历史全量）→ 提取所有 SKU 前缀（已知前缀优先匹配，未知取「品牌-颜色」）
    2. 合并当前已配置的值（有就预填，没有就留空 + 状态=待确认）
    3. 输出一张带校验的 xlsx

用法：
    python 生成商品配置模板.py
    python 生成商品配置模板.py --input-dir /data/data/workflow003/input
"""
import argparse
import glob
import os
import sys

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

BASE = os.path.dirname(os.path.abspath(__file__))
PROJ = os.path.dirname(BASE)
sys.path.insert(0, PROJ)
import paths  # noqa: E402

DEFAULT_INPUT = next((p for p in ("/data/data/workflow003/input",
                                  "/Users/zuo/n8n/files/data/workflow003/input")
                      if os.path.isdir(p)), "/data/data/workflow003/input")

# ============================================================================
# 列定义：(程序键, 表头, 宽度, 是否业务填写, 是否必填)
# ============================================================================
COLS = [
    ("sku",      "SKU前缀",          16, False, False),
    ("orders",   "历史订单数",        10, False, False),
    ("lines",    "历史行数",           9, False, False),
    ("known",    "当前是否已配置",     12, False, False),
    ("buy",      "采购价_默认",       12, True,  True),
    ("buy_rm",   "采购价_美国偏远",   14, True,  False),
    ("dom",      "国内运费_默认",     12, True,  True),
    ("dom_rm",   "国内运费_美国偏远", 14, True,  False),
    ("wt",       "重量(KG)",          10, True,  True),
    ("svc_us",   "服务费_美国",       11, True,  False),
    ("svc",      "服务费_其他",       11, True,  True),
    ("ov_c",     "运价覆盖_国家",     13, True,  False),
    ("ov_p",     "运价覆盖_公斤价",   13, True,  False),
    ("ov_f",     "运价覆盖_挂号费",   13, True,  False),
    ("fixed",    "账单固定金额$",     13, True,  False),
    ("pack",     "每计费单位件数",    13, True,  True),
    ("status",   "状态",             10, True,  True),
    ("by",       "填写人",           10, True,  False),
    ("note",     "备注",             40, True,  False),
]

HEAD_BG = "FFB26223"      # 只读列（深橙棕）
HEAD_FILL_BG = "FFF5B482"  # 填写列（浅橙）
MISSING_BG = "FFFFC7CE"    # 必填未填（淡红）
KNOWN_BG = "FFE4EFE4"      # 已配置行（淡绿）


def load_current_config():
    """当前配置：优先 data/*.json，其次 示例数据/1-商品配置.xlsx"""
    out = {}
    try:
        import calc
        cfg = calc.load_cfg(paths.data_dir())
        out = dict(cfg["products"])
        print(f"当前配置取自 data/：{len(out)} 个前缀")
        return out
    except Exception:
        pass
    f = os.path.join(PROJ, "示例数据", "1-商品配置.xlsx")
    if os.path.exists(f):
        df = pd.read_excel(f, dtype=object)
        df.columns = [str(c).strip() for c in df.columns]
        for _, r in df.iterrows():
            p = str(r.get("SKU前缀") or "").strip()
            if p:
                out[p] = r.to_dict()
        print(f"当前配置取自 {os.path.basename(f)}：{len(out)} 个前缀")
    return out


def scan_orders(input_dir):
    """扫所有订单导出 → (前缀 → 统计, 唯一前缀列表)"""
    files = sorted(glob.glob(os.path.join(input_dir, "订单导出数据_*.xlsx")),
                   key=os.path.getmtime)
    if not files:
        raise SystemExit(f"{input_dir} 里没有 订单导出数据_*.xlsx")
    frames = []
    for f in files:
        try:
            d = pd.read_excel(f, dtype=str, usecols=["SKU", "交易编号"])
        except Exception:
            continue
        d = d[d["SKU"].notna()]
        if len(d):
            frames.append(d)
    if not frames:
        raise SystemExit("订单文件里读不到 SKU 列")
    df = pd.concat(frames, ignore_index=True)
    return df, files


def make_prefixer(known):
    known_sorted = sorted(known, key=len, reverse=True)

    def f(sku):
        s = str(sku).strip()
        for p in known_sorted:                    # ① 已知前缀优先（最长匹配）
            if s == p or s.startswith(p + "-"):
                return p
        head = s.split("_")[0]                    # ② 未知：取「品牌-颜色」两段
        parts = head.split("-")
        return "-".join(parts[:2]) if len(parts) >= 2 else head
    return f


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input-dir", default=DEFAULT_INPUT)
    ap.add_argument("--out", default=os.path.join(PROJ, "输出", "商品配置-待填模板.xlsx"))
    args = ap.parse_args()

    cfg = load_current_config()
    df, files = scan_orders(args.input_dir)
    print(f"订单文件 {len(files)} 个，{len(df)} 行，唯一 SKU {df['SKU'].nunique()}")

    # 已知前缀既包含别名（cfg 已展开），也包含主前缀
    known_all = set(cfg) | {str((v or {}).get("SKU前缀") or "").strip() for v in cfg.values()}
    known_all.discard("")
    df = df.assign(前缀=df["SKU"].map(make_prefixer(known_all)))
    stat = (df.groupby("前缀")
              .agg(订单数=("交易编号", "nunique"), 行数=("SKU", "size"))
              .sort_values("订单数", ascending=False))

    # 配置里一行可能写多个前缀（逗号分隔）→ 模板按「主前缀」合并成一行
    def canon(p):
        row = cfg.get(p) or {}
        return str(row.get("主前缀") or row.get("SKU前缀") or p).split(",")[0].strip()

    stat_canon = {}
    for p, r in stat.iterrows():
        k = canon(p)
        cur = stat_canon.setdefault(k, {"订单数": 0, "行数": 0})
        cur["订单数"] += int(r["订单数"])
        cur["行数"] += int(r["行数"])
    cfg_canon = {}
    for p, row in cfg.items():
        cfg_canon.setdefault(canon(p), row)

    all_prefix = sorted(set(stat_canon) | set(cfg_canon),
                        key=lambda p: (-stat_canon.get(p, {}).get("订单数", 0), p))
    print(f"模板行数 {len(all_prefix)}（订单里 {len(stat)} + 仅在配置里 "
          f"{len(set(cfg) - set(stat))}）")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "商品配置"

    thin = Side(style="thin", color="FFB0B0B0")
    BD = Border(left=thin, right=thin, top=thin, bottom=thin)

    for i, (_, title, width, editable, _req) in enumerate(COLS, 1):
        c = ws.cell(row=1, column=i, value=title)
        c.font = Font(name="宋体", size=11, bold=True)
        c.fill = PatternFill("solid", fgColor=HEAD_FILL_BG if editable else HEAD_BG)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BD
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[1].height = 42

    def val(prov, key):
        v = prov.get(key)
        if v is None or (isinstance(v, float) and v != v):
            return None
        try:
            f = float(v)
            return int(f) if f == int(f) else f
        except (TypeError, ValueError):
            return str(v).strip() or None

    for n, p in enumerate(all_prefix, start=2):
        cur = cfg_canon.get(p) or {}
        known = bool(cur)
        shown = str(cur.get("SKU前缀") or p).strip() if known else p
        row = {
            "sku": shown,
            "orders": int(stat_canon.get(p, {}).get("订单数", 0)),
            "lines": int(stat_canon.get(p, {}).get("行数", 0)),
            "known": "是" if known else "否",
            "buy": val(cur, "采购价"), "buy_rm": val(cur, "偏远采购价"),
            "dom": val(cur, "国内运费"), "dom_rm": val(cur, "偏远国内运费"),
            "wt": val(cur, "重量"),
            "svc_us": val(cur, "服务费_美国"), "svc": val(cur, "服务费") or 1.15,
            "ov_c": val(cur, "运价覆盖_国家"),
            "ov_p": val(cur, "运价覆盖_公斤价"), "ov_f": val(cur, "运价覆盖_挂号费"),
            "fixed": val(cur, "账单固定金额"),
            "pack": val(cur, "每计费单位件数") or 1,
            "status": "已确认" if known else "待确认",
            "by": None, "note": None,
        }
        for i, (key, _t, _w, _e, _r) in enumerate(COLS, 1):
            c = ws.cell(row=n, column=i, value=row[key])
            c.font = Font(name="宋体", size=11)
            c.border = BD
            c.alignment = Alignment(horizontal="center", vertical="center",
                                    wrap_text=(key == "note"))
            if key in ("buy", "buy_rm", "dom", "dom_rm", "wt", "svc", "svc_us",
                       "ov_p", "ov_f", "fixed"):
                c.number_format = "0.00###"
            if key == "pack":
                c.number_format = "0"
        if known:
            for i in (1, 2, 3, 4):
                ws.cell(row=n, column=i).fill = PatternFill("solid", fgColor=KNOWN_BG)

    last = len(all_prefix) + 1
    ws.freeze_panes = "E2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{last}"

    # ---- 数据校验 ----
    idx = {k: i + 1 for i, (k, *_rest) in enumerate(COLS)}
    L = get_column_letter
    dvs = [
        (DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                        allow_blank=True, error="必须 ≥ 0"), f"{L(idx['buy'])}2:{L(idx['buy'])}{last}"),
        (DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                        allow_blank=True, error="必须 ≥ 0"), f"{L(idx['dom'])}2:{L(idx['dom'])}{last}"),
        (DataValidation(type="decimal", operator="between", formula1="0", formula2="20",
                        allow_blank=True, error="重量按 KG 填，0~20"), f"{L(idx['wt'])}2:{L(idx['wt'])}{last}"),
        (DataValidation(type="decimal", operator="between", formula1="1", formula2="2",
                        allow_blank=True, error="服务费应在 1~2"), f"{L(idx['svc'])}2:{L(idx['svc'])}{last}"),
        (DataValidation(type="whole", operator="greaterThanOrEqual", formula1="1",
                        allow_blank=True, error="每计费单位件数必须是 ≥1 的整数"),
         f"{L(idx['pack'])}2:{L(idx['pack'])}{last}"),
        (DataValidation(type="list", formula1='"已确认,待确认"', allow_blank=False,
                        error="只能填 已确认 / 待确认"), f"{L(idx['status'])}2:{L(idx['status'])}{last}"),
        (DataValidation(type="list", formula1='"美国,美国（偏远）,澳大利亚（1区）,澳大利亚（2区）,澳大利亚（3区）"',
                        allow_blank=True, error="运价覆盖一般指个别国家，可留空"),
         f"{L(idx['ov_c'])}2:{L(idx['ov_c'])}{last}"),
    ]
    for dv, rng in dvs:
        ws.add_data_validation(dv)
        dv.add(rng)

    # ---- 必填未填标红 ----
    req = [k for k, _t, _w, _e, r in COLS if r]
    for key in req:
        col = L(idx[key])
        ws.conditional_formatting.add(
            f"{col}2:{col}{last}",
            openpyxl.formatting.rule.CellIsRule(
                operator="equal", formula=['""'],
                fill=PatternFill("solid", fgColor=MISSING_BG)))

    # ---- 说明 sheet ----
    d = wb.create_sheet("填写说明")
    d.append(["字段", "是否必填", "含义 / 规则"])
    for key, title, _w, editable, r in COLS:
        if not editable:
            continue
        desc = {
            "buy": "所有国家（除美国偏远）的单件采购价，数字",
            "buy_rm": "只用于「美国（偏远）」单，留空 = 沿用「采购价_默认」",
            "dom": "单件国内运费，数字（可以填 0）",
            "dom_rm": "只用于「美国（偏远）」单，留空 = 沿用「国内运费_默认」",
            "wt": "单件重量，单位 KG，保留 2~3 位小数",
            "svc_us": "美国单的服务费（毛利），留空 = 用「服务费_其他」",
            "svc": "其他国家的服务费，通常 1.15；个别合同价可不同（如 1.121）",
            "ov_c": "仅当该 SKU 有特殊合同运价时填国家名，否则留空",
            "ov_p": "配套「运价覆盖_国家」：固定的公斤价",
            "ov_f": "配套「运价覆盖_国家」：固定的挂号费",
            "fixed": "仅当该 SKU 账单锁死金额时填（美元）；留空 = 正常计算",
            "pack": "几件算一个计费单位：1=一件一份；2=两件一份、三件两份；3=三件一份",
            "status": "填完选「已确认」；还没核实就留「待确认」",
            "by": "填你的名字，便于追溯",
            "note": "补充说明，程序不读",
        }.get(key, "")
        d.append([title, "必填" if r else "选填", desc])
    d.append([])
    d.append(["颜色说明", "", "深橙棕表头 = 只读（程序生成）；浅橙表头 = 需要你填写"])
    d.append(["", "", "淡绿 = 该 SKU 已有配置（预填了现值，核对后可改）；淡红 = 必填但还没填"])
    d.append([])
    d.append(["填完怎么交", "", "把本文件发回，或直接按这张表在飞书多维表格里维护"])
    for i, w in enumerate([18, 10, 96], 1):
        d.column_dimensions[get_column_letter(i)].width = w
    for c in d[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="FFD9D9D9")
    d.freeze_panes = "A2"

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    wb.save(args.out)
    print(f"已生成 {args.out}")
    print(f"  共 {len(all_prefix)} 行，其中未配置需新填 "
          f"{sum(1 for p in all_prefix if p not in cfg_canon)} 行")


if __name__ == "__main__":
    main()
