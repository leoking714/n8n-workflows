#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ROLE: 账务层（账单期 / 挂账台账 / 报价历史）读写
ENTRY:
  preview:  python3 billing_ledger.py preview --period <期号> [--include 1]
            → 返回历史挂账合计（含明细），供计算表底部使用
  commit:   python3 billing_ledger.py commit --period <期号> --detail <bill_detail.json>
            [--suspend "4207,4210"] [--calc <xlsx>] [--pdf <pdf>] [--include 1]
            → 写账单期 / 新增挂账 / 标记已并入 / 写报价历史
CALLED-BY: workflow005 生成计算表之后（手工或 n8n 节点）
DEPS: feishu_newbase.json（飞书凭据）

设计：
  挂账台账是唯一欠款来源。状态机：挂起 → 已并入 → 已收 / 作废
  「已并入」后下期不再重复带出（幂等：挂账ID = 期号-交易编号）
"""
import argparse
import json
import os
import sys
import urllib.request

CRED = "/data/data/workflow005/config/feishu_newbase.json"
T_PERIOD = "tblmsJLJI7kNBNel"    # 账单期
T_LEDGER = "tbltUKJgJrWaoSfP"    # 挂账台账
T_PRICE = "tblFf4R1orrOgCup"     # 报价历史
UHF = 2.3


def api(cfg, token, path, body=None, method=None):
    url = "https://open.feishu.cn/open-apis" + path
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, method=method,
                                 headers={"Content-Type": "application/json",
                                          "Authorization": "Bearer " + token})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode())


def token_of(cfg):
    req = urllib.request.Request(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        data=json.dumps({"app_id": cfg["app_id"], "app_secret": cfg["app_secret"]}).encode(),
        headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.loads(r.read().decode())["tenant_access_token"]


def all_records(cfg, token, tid):
    out, page = [], ""
    while True:
        q = "?page_size=500" + (f"&page_token={page}" if page else "")
        r = api(cfg, token, f"/bitable/v1/apps/{cfg['base']}/tables/{tid}/records/search{q}",
                {"filter": {"conjunction": "and", "conditions": []}})
        out += r.get("data", {}).get("items", [])
        if not r.get("data", {}).get("has_more"):
            break
        page = r["data"]["page_token"]
    return out


def flat(v):
    if isinstance(v, list):
        return "".join(str(x.get("text", "")) if isinstance(x, dict) else str(x) for x in v)
    if isinstance(v, dict):
        return str(v.get("text") or v.get("value") or "")
    return "" if v is None else str(v)


def load_orders(detail_path):
    d = json.load(open(detail_path, encoding="utf-8"))
    fx = float((d.get("params") or {}).get("FX") or 6.6)
    orders = {}
    for r in d.get("detail") or []:
        t = str(r.get("交易编号"))
        o = orders.setdefault(t, {"txn": t, "件数": 0, "SKU": set(), "行数": 0,
                                  "成本": 0.0, "运费": 0.0, "国内": 0.0,
                                  "欧盟税": 0.0, "账单": 0.0})
        sku = str(r.get("SKU") or "")
        if not sku.upper().startswith("SHIP"):
            o["件数"] += int(float(r.get("件数") or 0))
        o["SKU"].add(sku.split("_")[0])
        o["行数"] += 1
        for k, src in (("成本", "成本"), ("运费", "运费"), ("国内", "国内运费合计"),
                       ("欧盟税", "欧盟税")):
            o[k] += float(r.get(src) or 0)
        if r.get("账单金额") is not None:
            o["账单"] = float(r.get("账单金额"))
    for o in orders.values():
        o["收入USD"] = o["账单"] + (UHF if o["欧盟税"] > 0 else 0)
        o["毛利RMB"] = o["收入USD"] * fx - o["成本"] - o["运费"] - o["国内"]
        o["毛利率"] = (o["毛利RMB"] / (o["收入USD"] * fx)) if o["收入USD"] else 0.0
    return orders, fx


def find_cell(items, key, value):
    for it in items:
        if flat(it["fields"].get(key)).strip() == str(value).strip():
            return it
    return None


def preview(args):
    cfg = json.load(open(CRED, encoding="utf-8"))
    token = token_of(cfg)
    rows = all_records(cfg, token, T_LEDGER)
    pending = []
    for it in rows:
        f = it["fields"]
        if flat(f.get("状态")).strip() != "挂起":
            continue
        src = flat(f.get("来源期")).strip()
        if args.include and src == args.period:
            continue          # 本期的挂起不重复带入
        pending.append({
            "挂账ID": flat(f.get("挂账ID")),
            "交易编号": flat(f.get("交易编号")),
            "来源期": src,
            "金额USD": float(f.get("金额USD") or 0),
        })
    print(json.dumps({
        "ok": True,
        "period": args.period,
        "历史挂账USD": round(sum(p["金额USD"] for p in pending), 2),
        "笔数": len(pending),
        "明细": pending,
    }, ensure_ascii=False))
    return pending


def commit(args):
    cfg = json.load(open(CRED, encoding="utf-8"))
    token = token_of(cfg)
    orders, fx = load_orders(args.detail)
    suspend = [s.strip() for s in (args.suspend or "").split(",") if s.strip()]

    # 1) 历史挂账：把「挂起」且非本期 的标记为「已并入」
    if args.include:
        rows = all_records(cfg, token, T_LEDGER)
        ups = []
        for it in rows:
            f = it["fields"]
            if flat(f.get("状态")).strip() == "挂起" and flat(f.get("来源期")).strip() != args.period:
                ups.append({"record_id": it["record_id"],
                            "fields": {"状态": "已并入", "并入期": args.period}})
        if ups:
            api(cfg, token, f"/bitable/v1/apps/{cfg['base']}/tables/{T_LEDGER}/records/batch_update",
                {"records": ups})

    # 2) 本期挂起：新增挂账记录
    adds = []
    for t in suspend:
        o = orders.get(t)
        if not o:
            continue
        adds.append({"挂账ID": f"{args.period}-{t}", "交易编号": t, "来源期": args.period,
                     "金额USD": round(o["收入USD"], 2), "状态": "挂起", "备注": args.note or ""})
    if adds:
        api(cfg, token, f"/bitable/v1/apps/{cfg['base']}/tables/{T_LEDGER}/records/batch_create",
            {"records": [{"fields": x} for x in adds]})

    # 3) 账单期：upsert
    orders_n = len(orders)
    qty = sum(o["件数"] for o in orders.values())
    income = sum(o["收入USD"] for o in orders.values())
    cost = sum(o["成本"] + o["运费"] + o["国内"] for o in orders.values())
    gp = income * fx - cost
    period_fields = {
        "期号": args.period, "开始": args.start or "", "结束": args.end or "",
        "订单数": orders_n, "件数": qty,
        "收入USD": round(income, 2), "成本RMB": round(cost, 2),
        "毛利RMB": round(gp, 2), "毛利率": (gp / (income * fx)) if income else 0,
        "状态": "草稿", "计算表": args.calc or "", "PDF": args.pdf or "", "备注": args.note or "",
    }
    exist = find_cell(all_records(cfg, token, T_PERIOD), "期号", args.period)
    if exist:
        api(cfg, token, f"/bitable/v1/apps/{cfg['base']}/tables/{T_PERIOD}/records/batch_update",
            {"records": [{"record_id": exist["record_id"], "fields": period_fields}]})
        period_action = "update"
    else:
        api(cfg, token, f"/bitable/v1/apps/{cfg['base']}/tables/{T_PERIOD}/records/batch_create",
            {"records": [{"fields": period_fields}]})
        period_action = "create"

    # 4) 报价历史：按 (账期, SKU, 件数) 去重写入
    hist_rows = all_records(cfg, token, T_PRICE)
    seen = {(flat(r["fields"].get("账期")), flat(r["fields"].get("SKU")), flat(r["fields"].get("件数")))
            for r in hist_rows}
    recs = []
    for o in orders.values():
        if not o["收入USD"] or not o["件数"]:
            continue
        price = round(o["收入USD"] / o["件数"], 2)
        for sku in sorted(o["SKU"]):
            key = (args.period, sku, str(o["件数"]))
            if key in seen:
                continue
            seen.add(key)
            recs.append({"fields": {"账期": args.period, "SKU": sku, "件数": o["件数"],
                                    "单价USD": price, "毛利率": round(o["毛利率"], 4),
                                    "交易编号": o["txn"]}})
    for i in range(0, len(recs), 500):
        api(cfg, token, f"/bitable/v1/apps/{cfg['base']}/tables/{T_PRICE}/records/batch_create",
            {"records": recs[i:i + 500]})

    print(json.dumps({
        "ok": True, "period": args.period, "账单期": period_action,
        "挂账新增": len(adds), "报价历史新增": len(recs),
        "订单数": orders_n, "件数": qty, "收入USD": round(income, 2),
        "毛利RMB": round(gp, 2), "挂起明细": [a["交易编号"] for a in adds],
    }, ensure_ascii=False))


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    p1 = sub.add_parser("preview")
    p1.add_argument("--period", required=True)
    p1.add_argument("--include", type=int, default=1)
    p2 = sub.add_parser("commit")
    p2.add_argument("--period", required=True)
    p2.add_argument("--detail", required=True)
    p2.add_argument("--suspend", default="")
    p2.add_argument("--calc", default="")
    p2.add_argument("--pdf", default="")
    p2.add_argument("--start", default="")
    p2.add_argument("--end", default="")
    p2.add_argument("--note", default="")
    p2.add_argument("--include", type=int, default=1)
    a = ap.parse_args()
    if a.cmd == "preview":
        preview(a)
    else:
        commit(a)


if __name__ == "__main__":
    main()
