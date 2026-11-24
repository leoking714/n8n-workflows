#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ROLE: 毛利分析报表（单品/多品、历史报价对比、与上一份账单对比）
ENTRY: python3 margin_report.py --detail <bill_detail.json> [--prev <prev.json>] [--out <xlsx>]
CALLED-BY: 手工 / 后续接入 workflow005 生成计算表之后
DEPS: openpyxl

口径（全部分母均为「本期账单金额」）：
  收入USD      = Σ账单金额 + 欧盟单×UHF(2.3)          ← 与 PDF 的 Amount 一致
  产品成本RMB  = Σ成本
  物流成本RMB  = Σ运费（含挂号费）
  国内运费RMB  = Σ国内运费合计
  税费加收USD  = Σ欧盟税 + 欧盟单×UHF
  附加费加收RMB= Σ(附加费×服务费)
  毛利RMB      = 收入USD×汇率 − (产品成本+物流成本+国内运费)
  毛利率       = 毛利RMB / (收入USD×汇率)
"""
import argparse
import glob
import json
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DETAIL_DIR = "/data/data/workflow005/logs"
UHF = 2.3

HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
HEAD_FONT = Font(bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(bold=True, size=11)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def num(x, d=0.0):
    try:
        return float(x)
    except Exception:
        return d


def load(path):
    d = json.load(open(path, encoding="utf-8"))
    rows = d.get("detail") or []
    fx = num((d.get("params") or {}).get("FX"), 6.6) or 6.6
    return d, rows, fx


def orders_of(rows, fx=6.6):
    """按交易编号聚合：成本 / 运费 / 国内运费 / 税费 / 附加费 / 件数 / SKU"""
    out = {}
    for r in rows:
        t = str(r.get("交易编号"))
        o = out.setdefault(t, {
            "txn": t, "件数": 0, "SKU": set(), "行数": 0,
            "成本": 0.0, "运费": 0.0, "国内": 0.0,
            "欧盟税": 0.0, "附加费": 0.0, "附加费x服务费": 0.0,
            "账单": None, "毛利RMB": 0.0,
        })
        sku = str(r.get("SKU") or "")
        qty = int(num(r.get("件数"), 0))
        is_ship = sku.upper().startswith("SHIP")
        if not is_ship:
            o["件数"] += qty
        o["SKU"].add(sku.split("_")[0])
        o["行数"] += 1
        o["成本"] += num(r.get("成本"))
        o["运费"] += num(r.get("运费"))
        o["国内"] += num(r.get("国内运费合计"))
        o["欧盟税"] += num(r.get("欧盟税"))
        o["附加费"] += num(r.get("附加费"))
        o["附加费x服务费"] += num(r.get("附加费")) * num(r.get("服务费"), 1)
        if r.get("账单金额") is not None:
            o["账单"] = num(r.get("账单金额"))
    for o in out.values():
        if o["账单"] is None:
            o["账单"] = 0.0
        o["收入USD"] = o["账单"] + (UHF if o["欧盟税"] > 0 else 0)
        o["收入RMB"] = o["收入USD"] * fx
        o["毛利RMB"] = o["收入RMB"] - o["成本"] - o["运费"] - o["国内"]
        o["毛利率"] = (o["毛利RMB"] / o["收入RMB"]) if o["收入RMB"] else 0.0
    return out


def summarize(orders, fx):
    s = {"订单数": len(orders), "件数": 0, "收入USD": 0.0, "产品成本": 0.0,
         "物流成本": 0.0, "国内运费": 0.0, "税费加收USD": 0.0, "附加费加收": 0.0}
    for o in orders.values():
        s["件数"] += o["件数"]
        s["收入USD"] += o["账单"] + (UHF if o["欧盟税"] > 0 else 0)
        s["产品成本"] += o["成本"]
        s["物流成本"] += o["运费"]
        s["国内运费"] += o["国内"]
        s["税费加收USD"] += o["欧盟税"] + (UHF if o["欧盟税"] > 0 else 0)
        s["附加费加收"] += o["附加费x服务费"]
    s["收入RMB"] = s["收入USD"] * fx
    s["毛利RMB"] = s["收入RMB"] - s["产品成本"] - s["物流成本"] - s["国内运费"]
    s["毛利率"] = (s["毛利RMB"] / s["收入RMB"]) if s["收入RMB"] else 0
    return s


def style_sheet(ws, widths, header_row=1):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for c in ws[header_row]:
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = ws.cell(header_row + 1, 1)


def sheet_summary(wb, cur, prev):
    ws = wb.create_sheet("汇总")
    ws.append(["指标", "本次账单", "上一份账单", "差异", "变化率"])
    rows = [
        ("订单数", "订单数", 0, "#,##0"),
        ("件数", "件数", 0, "#,##0"),
        ("收入 USD", "收入USD", 2, "#,##0.00"),
        ("产品成本 ￥", "产品成本", 2, "#,##0.00"),
        ("物流成本 ￥", "物流成本", 2, "#,##0.00"),
        ("国内运费 ￥", "国内运费", 2, "#,##0.00"),
        ("税费加收 USD", "税费加收USD", 2, "#,##0.00"),
        ("附加费加收 ￥", "附加费加收", 2, "#,##0.00"),
        ("毛利 ￥", "毛利RMB", 2, "#,##0.00"),
    ]
    for label, key, dec, fmt in rows:
        a = cur[key]
        b = (prev or {}).get(key)
        diff = (a - b) if isinstance(b, (int, float)) else None
        rate = (diff / b) if (diff is not None and b) else None
        ws.append([label, round(a, dec),
                   (round(b, dec) if isinstance(b, (int, float)) else "-"),
                   (round(diff, dec) if diff is not None else "-"),
                   (rate if rate is not None else "-")])
        r = ws.max_row
        for c in (2, 3, 4):
            ws.cell(r, c).number_format = fmt
        ws.cell(r, 5).number_format = "0.0%"
    a = cur["毛利率"]
    b = (prev or {}).get("毛利率")
    ws.append(["毛利率", a, (b if b is not None else "-"),
               ((a - b) if b is not None else "-"), "-"])
    r = ws.max_row
    for c in (2, 3, 4):
        ws.cell(r, c).number_format = "0.00%"

    # 口径说明
    ws.append([])
    ws.append(["口径说明", "", "", "", ""])
    for line in (
        "收入 = 账单金额 + 欧盟单 × 2.3(UHF)，与 PDF 的 Amount 一致",
        "产品成本 = Σ采购成本；物流成本 = Σ运费(含挂号费)；国内运费 = Σ国内运费",
        "税费加收 = Σ欧盟税 + 欧盟单 × 2.3；附加费加收 = Σ(附加费 × 服务费)",
        "毛利 ￥ = 收入 × 汇率 − (产品成本 + 物流成本 + 国内运费)",
        "上一份账单：自动取「订单集合不同且 ≥100 单」的最近一份明细；没有则留空",
    ):
        ws.append([line, "", "", "", ""])
    ws.cell(ws.max_row - 5, 1).font = TITLE_FONT
    style_sheet(ws, [16, 16, 16, 14, 12])
    return ws


def sheet_single_multi(wb, orders):
    ws = wb.create_sheet("单品vs多品")
    ws.append(["类型", "订单数", "件数", "收入 USD", "成本合计 ￥", "毛利 ￥", "毛利率"])
    groups = {"单品订单（1 个 SKU）": [], "多品订单（≥2 个 SKU）": []}
    for o in orders.values():
        groups["单品订单（1 个 SKU）" if len(o["SKU"]) <= 1 else "多品订单（≥2 个 SKU）"].append(o)
    for name, lst in groups.items():
        rev_usd = sum(o["收入USD"] for o in lst)
        rev_rmb = sum(o["收入RMB"] for o in lst)
        cost = sum(o["成本"] + o["运费"] + o["国内"] for o in lst)
        gp = rev_rmb - cost
        ws.append([name, len(lst), sum(o["件数"] for o in lst), round(rev_usd, 2),
                   round(cost, 2), round(gp, 2), (gp / rev_rmb if rev_rmb else 0)])
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 4).number_format = "#,##0.00"
        ws.cell(r, 5).number_format = "#,##0.00"
        ws.cell(r, 6).number_format = "#,##0.00"
        ws.cell(r, 7).number_format = "0.00%"
    style_sheet(ws, [22, 10, 10, 14, 14, 14, 10])
    return ws


def history_index(exclude_path, cur_txns):
    """历史报价索引：(SKU, 件数) → 最近一次（只取「不同账期」的明细：订单重叠 ≤ 50%）"""
    idx = {}
    cur = set(cur_txns)
    for p in sorted(glob.glob(os.path.join(DETAIL_DIR, "bill_detail_*.json")), key=os.path.getmtime):
        if os.path.abspath(p) == os.path.abspath(exclude_path):
            continue
        try:
            d, rows, fx = load(p)
        except Exception:
            continue
        ptx = set(orders_of(rows, fx))
        if ptx and len(ptx & cur) * 2 > len(ptx):
            continue   # 同一账期的重复跑 → 不算历史
        for o in orders_of(rows, fx).values():
            if not o["收入USD"]:
                continue
            for sku in o["SKU"]:
                key = (sku, o["件数"])
                idx[key] = {  # 后面的覆盖前面的 → 得到「最近一次」
                    "报价": o["账单"] / o["件数"] if o["件数"] else 0,
                    "毛利率": o["毛利率"],
                    "账单": o["账单"],
                    "文件": os.path.basename(p),
                    "日期": os.path.getmtime(p),
                }
    return idx


def sheet_compare(wb, orders, hist):
    ws = wb.create_sheet("同款历史对比")
    ws.append(["SKU", "件数", "本次报价(USD/件)", "本次毛利率", "上次报价(USD/件)",
               "上次毛利率", "报价变化", "毛利率变化", "上次来源"])
    for o in orders.values():
        if not o["收入USD"]:
            continue
        for sku in sorted(o["SKU"]):
            key = (sku, o["件数"])
            cur_price = o["账单"] / o["件数"] if o["件数"] else 0
            h = hist.get(key)
            if not h:
                ws.append([sku, o["件数"], round(cur_price, 2), o["毛利率"], "-", "-", "-", "-", "无历史"])
                continue
            ws.append([sku, o["件数"], round(cur_price, 2), o["毛利率"],
                       round(h["报价"], 2), h["毛利率"],
                       round(cur_price - h["报价"], 2), o["毛利率"] - h["毛利率"],
                       h["文件"]])
    for r in range(2, ws.max_row + 1):
        for c in (3, 5, 7):
            ws.cell(r, c).number_format = "#,##0.00"
        for c in (4, 6, 8):
            ws.cell(r, c).number_format = "0.00%"
    style_sheet(ws, [18, 8, 16, 12, 16, 12, 12, 12, 30])
    return ws


def sheet_orders(wb, orders):
    ws = wb.create_sheet("逐单毛利")
    ws.append(["交易编号", "SKU数", "件数", "收入 USD", "产品成本 ￥", "物流成本 ￥",
               "国内运费 ￥", "税费加收 USD", "毛利 ￥", "毛利率"])
    for o in sorted(orders.values(), key=lambda x: -x["毛利RMB"]):
        ws.append([o["txn"], len(o["SKU"]), o["件数"], round(o["收入USD"], 2),
                   round(o["成本"], 2), round(o["运费"], 2), round(o["国内"], 2),
                   round(o["欧盟税"] + (UHF if o["欧盟税"] > 0 else 0), 2),
                   round(o["毛利RMB"], 2), o["毛利率"]])
    for r in range(2, ws.max_row + 1):
        for c in (4, 5, 6, 7, 8, 9):
            ws.cell(r, c).number_format = "#,##0.00"
        ws.cell(r, 10).number_format = "0.00%"
    style_sheet(ws, [10, 8, 8, 12, 12, 12, 12, 12, 12, 10])
    return ws


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--detail", required=True)
    ap.add_argument("--prev", default="")
    ap.add_argument("--out", default="")
    a = ap.parse_args()

    d, rows, fx = load(a.detail)
    orders = orders_of(rows, fx)
    cur = summarize(orders, fx)

    prev_sum = None
    if a.prev:
        _, prows, pfx = load(a.prev)
        prev_sum = summarize(orders_of(prows, pfx), pfx)
    else:
        # 自动取上一份：最近一个「订单集合不同」的明细
        cur_txn = set(orders)
        for p in sorted(glob.glob(os.path.join(DETAIL_DIR, "bill_detail_*.json")),
                        key=os.path.getmtime, reverse=True):
            if os.path.abspath(p) == os.path.abspath(a.detail):
                continue
            try:
                _, prows, pfx = load(p)
            except Exception:
                continue
            porders = orders_of(prows, pfx)
            if (set(porders) != cur_txn and len(porders) >= 100
                    and sum(o["收入USD"] for o in porders.values()) > 0):
                prev_sum = summarize(porders, pfx)
                prev_sum["来源"] = os.path.basename(p)
                break

    wb = Workbook()
    wb.remove(wb.active)
    sheet_summary(wb, cur, prev_sum)
    sheet_single_multi(wb, orders)
    sheet_compare(wb, orders, history_index(a.detail, orders))
    sheet_orders(wb, orders)

    out = a.out or os.path.join(DETAIL_DIR, "毛利分析_" + os.path.basename(a.detail).replace("bill_detail_", "").replace(".json", "") + ".xlsx")
    wb.save(out)
    msg = {
        "ok": True, "out": out, "orders": cur["订单数"], "件数": cur["件数"],
        "收入USD": round(cur["收入USD"], 2), "毛利RMB": round(cur["毛利RMB"], 2),
        "毛利率": round(cur["毛利率"], 4),
        "上一份": (prev_sum or {}).get("来源", "无"),
    }
    print(json.dumps(msg, ensure_ascii=False))


if __name__ == "__main__":
    main()
