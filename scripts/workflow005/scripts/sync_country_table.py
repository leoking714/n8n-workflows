#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
按「运价-分段」里的国家，同步「国家」表：
  · 缺的国家自动加一行
  · 服务费：指定国家 1.15，其余默认 1.11
  · 欧盟国家：是否欧盟=是、欧盟税额$=5.6，其余 0
  · 订单附加费：默认 0（可手工再改）

用法：
  python sync_country_table.py [--dry-run]
"""
import argparse, json, sys, urllib.request

CRED = "/data/data/workflow005/config/feishu_newbase.json"
RATE_TABLE = "tblNYhByd3ivGP4L"
COUNTRY_TABLE = "tblnP5e6VWdU1j1T"

# 服务费 1.15 的国家（其余 1.11）
FEE_115 = {
    "澳大利亚（1区）", "澳大利亚（2区）", "澳大利亚（3区）",
    "加拿大", "卡塔尔", "科威特", "墨西哥", "瑞士", "新西兰",
    "以色列", "智利", "中国香港", "巴林",
}

# 欧盟国家（税额 5.6 美元）
EU = {
    "瑞典", "丹麦", "芬兰", "爱尔兰", "荷兰", "比利时", "德国", "法国",
    "意大利", "西班牙", "卢森堡", "奥地利", "希腊", "马耳他", "葡萄牙",
    "塞浦路斯", "波兰", "捷克", "斯洛伐克", "匈牙利", "斯洛文尼亚",
    "克罗地亚", "罗马尼亚", "保加利亚", "拉脱维亚", "立陶宛", "爱沙尼亚",
}
EU_TAX = 5.6
DEFAULT_FEE = 1.11


def api(path, token, body=None):
    url = "https://open.feishu.cn/open-apis" + path
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data,
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


def flat(v):
    if isinstance(v, list):
        return "".join(str(x.get("text", "")) if isinstance(x, dict) else str(x) for x in v)
    if isinstance(v, dict):
        return str(v.get("text") or v.get("value") or "")
    return "" if v is None else str(v)


def all_records(cfg, token, tid):
    out, page = [], ""
    while True:
        q = "?page_size=500" + (f"&page_token={page}" if page else "")
        r = api(f"/bitable/v1/apps/{cfg['base']}/tables/{tid}/records/search{q}", token,
                {"filter": {"conjunction": "and", "conditions": []}})
        out += r.get("data", {}).get("items", [])
        if not r.get("data", {}).get("has_more"):
            break
        page = r["data"]["page_token"]
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    cfg = json.load(open(CRED, encoding="utf-8"))
    token = token_of(cfg)

    rate_rows = all_records(cfg, token, RATE_TABLE)
    rate_countries = sorted({flat(r["fields"].get("国家")).strip() for r in rate_rows} - {""})
    cty_rows = all_records(cfg, token, COUNTRY_TABLE)
    existing = {}
    for r in cty_rows:
        c = flat(r["fields"].get("国家")).strip()
        if c:
            existing[c] = r

    def desired(c):
        return {
            "国家": c,
            "是否欧盟": "是" if c in EU else "否",
            "欧盟税额$": EU_TAX if c in EU else 0,
            "订单附加费": 0,
            "服务费覆盖": 1.15 if c in FEE_115 else DEFAULT_FEE,
        }

    adds, updates = [], []
    for c in rate_countries:
        want = desired(c)
        if c not in existing:
            adds.append(want)
            continue
        cur = existing[c]["fields"]
        cur_fee = flat(cur.get("服务费覆盖"))
        cur_eu = flat(cur.get("是否欧盟"))
        need = {}
        if str(cur_fee) != str(want["服务费覆盖"]):
            need["服务费覆盖"] = want["服务费覆盖"]
        if cur_eu != want["是否欧盟"]:
            need["是否欧盟"] = want["是否欧盟"]
            need["欧盟税额$"] = want["欧盟税额$"]
        if need:
            updates.append({"record_id": existing[c]["record_id"], "fields": need})

    print(f"运价里国家 {len(rate_countries)} 个 | 国家表现有 {len(existing)} 行")
    print(f"需新增 {len(adds)} | 需改服务费/欧盟标记 {len(updates)}")
    for x in adds:
        print("   + ", x["国家"], "| 服务费", x["服务费覆盖"], "| 欧盟", x["是否欧盟"])
    for u in updates[:10]:
        print("   ~ ", u["fields"])

    if a.dry_run:
        print(json.dumps({"ok": True, "dry_run": True, "add": len(adds), "update": len(updates)}, ensure_ascii=False))
        return

    if adds:
        for i in range(0, len(adds), 500):
            api(f"/bitable/v1/apps/{cfg['base']}/tables/{COUNTRY_TABLE}/records/batch_create", token,
                {"records": [{"fields": x} for x in adds[i:i + 500]]})
    if updates:
        for i in range(0, len(updates), 500):
            api(f"/bitable/v1/apps/{cfg['base']}/tables/{COUNTRY_TABLE}/records/batch_update", token,
                {"records": updates[i:i + 500]})

    print(json.dumps({"ok": True, "add": len(adds), "update": len(updates)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
