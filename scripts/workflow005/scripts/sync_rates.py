#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
商品表加了新前缀 → 自动在运价表补齐「该前缀 × 所有国家」的行，并建对应视图。

规则：
  · 商品表里每个 SKU前缀，在运价表应各有 len(国家) 行
  · 缺的行按「公斤价/挂号费留空」创建（等业务在视图里填价）
  · 每个前缀建一个视图「前缀·XXX」；另外建一个「待补价」视图（公斤价为空）

用法：
  python sync_rates.py            # 实际执行
  python sync_rates.py --dry-run  # 只看要做什么
"""
import argparse, json, os, sys, urllib.request

CFG = "/data/data/workflow005/config/feishu_newbase.json"


def post(url, token, body):
    req = urllib.request.Request(url, data=json.dumps(body).encode(),
                                 headers={"Content-Type": "application/json",
                                          "Authorization": "Bearer " + token})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode())


def get(url, token):
    req = urllib.request.Request(url, headers={"Authorization": "Bearer " + token})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode())


def flat(v):
    if v is None:
        return ""
    if isinstance(v, list):
        return "".join(str(x.get("text", "")) if isinstance(x, dict) else str(x) for x in v)
    if isinstance(v, dict):
        return str(v.get("text") or v.get("value") or "")
    return str(v)


def all_records(base, token, tid):
    items, page = [], ""
    while True:
        q = "?page_size=500" + (f"&page_token={page}" if page else "")
        r = post(f"https://open.feishu.cn/open-apis/bitable/v1/apps/{base}/tables/{tid}/records/search{q}",
                 token, {"filter": {"conjunction": "and", "conditions": []}})
        items += r.get("data", {}).get("items", [])
        if not r.get("data", {}).get("has_more"):
            break
        page = r["data"]["page_token"]
    return items


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    cfg = json.load(open(CFG, encoding="utf-8"))
    base, T = cfg["base"], cfg["tables"]

    tk = post("https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal", "",
              {"app_id": cfg["app_id"], "app_secret": cfg["app_secret"]})
    token = tk["tenant_access_token"]

    prods = all_records(base, token, T["商品"])
    rates = all_records(base, token, T["运价"])
    ctys = all_records(base, token, T["国家"])

    prefixes = sorted({flat(p["fields"].get("SKU前缀")).strip() for p in prods} - {""})
    prefixes = [p for p in prefixes if not p.startswith("(待补)")]
    countries = sorted({flat(c["fields"].get("国家")).strip() for c in ctys} - {""})
    have = set()
    for r in rates:
        f = r["fields"]
        have.add((flat(f.get("SKU前缀")).strip(), flat(f.get("国家")).strip()))

    need = [(p, c) for p in prefixes for c in countries if (p, c) not in have]
    print(f"商品前缀 {len(prefixes)} 个 | 国家 {len(countries)} 个 | 运价已有 {len(have)} 行 | 需补 {len(need)} 行")

    if need and not a.dry_run:
        for i in range(0, len(need), 500):
            batch = [{"fields": {"SKU前缀": p, "国家": c}} for p, c in need[i:i + 500]]
            r = post(f"https://open.feishu.cn/open-apis/bitable/v1/apps/{base}/tables/{T['运价']}/records/batch_create",
                     token, {"records": batch})
            if r.get("code") != 0:
                print("补行失败:", r.get("code"), r.get("msg"))

    # 视图
    views = get(f"https://open.feishu.cn/open-apis/bitable/v1/apps/{base}/tables/{T['运价']}/views?page_size=100", token)
    existing = {v["view_name"] for v in views.get("data", {}).get("items", [])}
    created = []
    def mk(name, field, value):
        r = post(f"https://open.feishu.cn/open-apis/bitable/v1/apps/{base}/tables/{T['运价']}/views", token,
                 {"view_name": name, "view_type": "grid",
                  "property": {"filter_info": {"conjunction": "and",
                                               "conditions": [{"field_name": field, "operator": "is",
                                                               "value": [str(value)]}]}}})
        return r.get("code") == 0

    for p in prefixes:
        name = "前缀·" + p
        if name in existing:
            continue
        if a.dry_run:
            created.append(name); continue
        if mk(name, "SKU前缀", p):
            created.append(name)

    if "待补价" not in existing:
        if a.dry_run:
            created.append("待补价")
        else:
            r = post(f"https://open.feishu.cn/open-apis/bitable/v1/apps/{base}/tables/{T['运价']}/views", token,
                     {"view_name": "待补价", "view_type": "grid",
                      "property": {"filter_info": {"conjunction": "and",
                                                   "conditions": [{"field_name": "公斤价", "operator": "isEmpty"}]}}})
            if r.get("code") == 0:
                created.append("待补价")

    print(json.dumps({"ok": True, "补行": len(need), "新建视图": created,
                      "dry_run": a.dry_run}, ensure_ascii=False))


if __name__ == "__main__":
    main()
