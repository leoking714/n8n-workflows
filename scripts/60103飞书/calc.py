# -*- coding: utf-8 -*-
"""
60103 计算核心（新工程，独立实现）

输入：
    data/*.json     飞书多维表格同步下来的五张配置表
    订单导出 xlsx    一行 = 一件商品

输出：
    detail DataFrame（每行一件商品，含全部金额列）
    issues 列表（阻断/警告）

口径与老脚本一致，数据源换成飞书。老脚本不受影响。
"""
import json
import os
import re

import numpy as np
import pandas as pd

import paths

BASE = os.path.dirname(os.path.abspath(__file__))

# ============================================================================
# 1. 读配置（飞书 JSON）
# ============================================================================
def load_cfg(data_dir=None):
    """读 data/*.json → 结构化配置字典"""
    d = data_dir or paths.data_dir()

    # 两种存法都支持：
    #   ① data/config.json 里一个对象装 5 张表（n8n Lark 节点版用这个）
    #   ② data/<表名>.json 各一个文件（sync_config.py 用这个）
    bundle_path = os.path.join(d, "config.json")
    bundle = None
    if os.path.exists(bundle_path):
        with open(bundle_path, encoding="utf-8") as f:
            bundle = json.load(f)

    def read(name):
        if bundle is not None:
            if name not in bundle:
                raise KeyError(f"config.json 里缺少「{name}」")
            return bundle[name]
        p = os.path.join(d, f"{name}.json")
        if not os.path.exists(p):
            raise FileNotFoundError(f"缺少配置 {p}，先跑 sync_config.py")
        with open(p, encoding="utf-8") as f:
            return json.load(f)

    raw_params = read("全局参数")
    params = {}
    for r in raw_params:
        k = str(r.get("参数键") or "").strip()
        if not k:
            continue
        v = r.get("参数值")
        try:
            params[k] = float(v)
        except (TypeError, ValueError):
            params[k] = v

    # 商品配置：前缀 → 行
    #   一个单元格里可以写多个前缀（英文逗号 / 中文逗号 / 顿号分隔），
    #   表示这些前缀是同一件商品（同价、同重、同服务费），展开成多条别名
    prods, alias_owner = {}, {}
    for r in read("商品配置"):
        cell = str(r.get("SKU前缀") or "").strip()
        if not cell:
            continue
        aliases = [a.strip() for a in re.split(r"[,，、]", cell) if a.strip()]
        if not aliases:
            continue
        row = dict(r)
        row["主前缀"] = aliases[0]              # 展示/异常里用第一个
        row["组"] = cell                        # 同一行的别名共享计费单位
        for a in aliases:
            if a in prods:
                raise ValueError(
                    f"商品配置里前缀「{a}」重复出现（分别在「{prods[a].get('主前缀')}」"
                    f"和「{row['主前缀']}」），同一前缀只能配一次")
            prods[a] = row

    # 运价表：国家 → [(下界, 上界, 公斤价, 挂号费)]
    rates, margins = {}, {}
    for r in read("运价表"):
        c = str(r.get("国家") or "").strip()
        lo, hi = r.get("下界"), r.get("上界")
        if not c or not isinstance(lo, (int, float)) or not isinstance(hi, (int, float)):
            continue
        rates.setdefault(c, []).append(
            (float(lo), float(hi), r.get("公斤价"), r.get("挂号费")))
        m = r.get("毛利")
        if c not in margins and isinstance(m, (int, float)):
            margins[c] = float(m)
    for c in rates:
        rates[c].sort(key=lambda x: x[0])

    # 澳洲分区：邮编 → 分区
    au = {}
    for r in read("澳洲分区"):
        p = str(r.get("邮编") or "").strip()
        z = r.get("分区")
        if len(p) == 4 and isinstance(z, (int, float)):
            au.setdefault(p, int(z))

    # 国家规则：国家 → 规则
    rules = {}
    for r in read("国家规则"):
        c = str(r.get("国家") or "").strip()
        if not c or str(r.get("状态") or "生效") != "生效":
            continue
        rules[c] = {
            "is_eu": str(r.get("是否欧盟") or "否").strip() == "是",
            "surcharge": float(r.get("订单附加费") or 0),
            "service": r.get("服务费覆盖"),
        }

    return {"params": params, "products": prods, "rates": rates,
            "margins": margins, "au": au, "country": rules}


def P(cfg, key, default=None):
    v = cfg["params"].get(key)
    return float(v) if isinstance(v, (int, float)) else default


# ============================================================================
# 2. 读订单
# ============================================================================
ORDER_COL_ALIAS = {"国家中文": "国家", "国家(中)": "国家", "邮政编码": "邮编"}
ORDER_FILL = ("交易编号", "订单编号", "国家", "物流公司", "物流渠道", "内部单号", "邮编")


def read_orders(path):
    """订单导出 xlsx → DataFrame，一行一件商品"""
    df = pd.read_excel(path, dtype=str)
    return normalize_orders(df)


def read_orders_json(path):
    """ERP 拉下来的订单 JSON → DataFrame（字段名与导出表一致）"""
    with open(path, encoding="utf-8") as f:
        rows = json.load(f)
    if isinstance(rows, dict):
        rows = rows.get("data") or rows.get("items") or rows.get("records") or []
    df = pd.DataFrame(rows)
    df = df.astype(str)
    return normalize_orders(df)


def normalize_orders(df):
    """订单表清洗：列名统一、订单级字段 ffill、件数转 int、标记首行"""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns={k: v for k, v in ORDER_COL_ALIAS.items() if k in df.columns})

    need = ["交易编号", "SKU"]
    miss = [c for c in need if c not in df.columns]
    if miss:
        raise ValueError(f"订单表缺少列 {miss}；实际列 = {list(df.columns)}")

    if "SKU" not in df.columns:
        raise ValueError("订单表缺少 SKU 列")

    df = df[df["SKU"].notna()].copy()
    for c in ORDER_FILL:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].ffill().fillna("")

    df["商品数量"] = (pd.to_numeric(df.get("商品数量"), errors="coerce")
                      .fillna(1).astype(int))
    df["源表行号"] = range(1, len(df) + 1)
    df["首行"] = df.groupby("交易编号", sort=False).cumcount() == 0
    return df.reset_index(drop=True)


# ============================================================================
# 3. 国家改写
# ============================================================================
def fix_country(df, cfg):
    """美国+岛屿 → 美国（偏远）；澳大利亚 → 澳大利亚（X区）"""
    df = df.copy()
    df["国家"] = df["国家"].fillna("").astype(str).str.strip()

    remote = df["国家"].eq("美国") & df["物流渠道"].astype(str).str.contains("岛屿", na=False)
    df.loc[remote, "国家"] = "美国（偏远）"
    df["偏远"] = df["国家"].eq("美国（偏远）")

    au = df["国家"].eq("澳大利亚")
    if au.any():
        zone = df.loc[au, "邮编"].astype(str).str.strip().str.zfill(4).map(cfg["au"])
        df.loc[au, "国家"] = zone.map(
            lambda z: f"澳大利亚（{int(z)}区）" if pd.notna(z) else "澳大利亚")
    return df


# ============================================================================
# 4. SKU 匹配
# ============================================================================
def match_sku(df, cfg):
    """最长前缀匹配 → 采购价/国内运费/重量/服务费/每计费件数/运价覆盖/固定金额"""
    df = df.copy()
    prefixes = sorted(cfg["products"], key=len, reverse=True)
    prods = cfg["products"]

    def find(sku):
        s = str(sku)
        for p in prefixes:
            if s == p or s.startswith(p + "-") or s.startswith(p + "_"):
                return p
        return None

    df["前缀"] = df["SKU"].map(find)
    # 同一配置行的多个别名（如 FUDGE-RST,FUDGE-BLK）共享计费单位
    df["配置组"] = df["前缀"].map(lambda p: (prods.get(p) or {}).get("组") or p)
    cols = ["采购价", "偏远采购价", "国内运费", "偏远国内运费", "重量", "服务费",
            "服务费_美国", "每计费件数", "运价覆盖_国家", "运价覆盖_公斤价",
            "运价覆盖_挂号费", "账单固定金额"]
    for c in cols:
        df[c] = df["前缀"].map(lambda p, c=c: (prods.get(p) or {}).get(c))

    for c in ("采购价", "偏远采购价", "国内运费", "偏远国内运费", "重量",
              "服务费", "服务费_美国", "每计费件数", "运价覆盖_公斤价",
              "运价覆盖_挂号费", "账单固定金额"):
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # 偏远价：偏远列有值就用，否则用普通列
    df["采购价"] = df["偏远采购价"].where(df["偏远"] & df["偏远采购价"].notna(), df["采购价"])
    df["国内运费"] = df["偏远国内运费"].where(df["偏远"] & df["偏远国内运费"].notna(), df["国内运费"])

    df["价格缺失"] = df["前缀"].notna() & df["采购价"].isna()

    # 兜底
    df["采购价"] = df["采购价"].fillna(P(cfg, "BUY_FALLBACK", 15))
    df["国内运费"] = df["国内运费"].fillna(P(cfg, "DOMESTIC_FALLBACK", 4))
    df["重量"] = df["重量"].fillna(P(cfg, "WEIGHT_FALLBACK", 0.17))
    df["每计费件数"] = df["每计费件数"].fillna(1).clip(lower=1)

    # 非实物行：SKU 以 SHIP 开头 且 没有商品图
    prefix_kw = str(P(cfg, "NON_GOODS_PREFIX", "SHIP") or "SHIP").upper()
    no_img = df.get("原始图片", pd.Series("", index=df.index)).fillna("").astype(str).str.strip().eq("")
    df["非实物"] = df["SKU"].astype(str).str.strip().str.upper().str.startswith(prefix_kw) & no_img
    df.loc[df["非实物"], ["采购价", "国内运费", "重量"]] = 0.0
    df.loc[df["非实物"], "价格缺失"] = False
    return df


# ============================================================================
# 5. 运价匹配（逐行）
# ============================================================================
def price_of(cfg, country, weight):
    """整单重量 → (公斤价, 挂号费)；区间上下界都包含；不兜底"""
    arr = cfg["rates"].get(country)
    if not arr:
        return None, None
    for lo, hi, p, fee in arr:
        if lo <= weight <= hi:      # 上下界都包含
            return p, fee
    return None, None


def match_rate(df, cfg):
    df = df.copy()
    df["行重量"] = df["商品数量"] * df["重量"]
    df["订单重量"] = (df.groupby("交易编号", sort=False)["行重量"]
                      .transform("sum").round(3))

    heads = df.loc[df["首行"], ["交易编号", "国家", "订单重量"]].drop_duplicates("交易编号")
    pairs = heads.apply(lambda r: price_of(cfg, r["国家"], r["订单重量"]), axis=1)
    heads["段公斤价"] = [x[0] for x in pairs]
    heads["段挂号费"] = [x[1] for x in pairs]

    n = len(df)
    df = df.merge(heads[["交易编号", "段公斤价", "段挂号费"]], on="交易编号", how="left")
    assert len(df) == n, "merge 后行数变了"

    use = (df["运价覆盖_国家"].notna()
           & df["运价覆盖_国家"].astype(str).eq(df["国家"])
           & df["运价覆盖_公斤价"].notna())
    df["公斤价"] = df["运价覆盖_公斤价"].where(use, df["段公斤价"])
    df["挂号费"] = df["运价覆盖_挂号费"].where(
        use & df["运价覆盖_挂号费"].notna(), df["段挂号费"])

    need = (~df["账单固定金额"].notna()) & (df["行重量"] > 0)
    bad = need & df["公斤价"].isna()
    df["缺运价"] = bad.groupby(df["交易编号"]).transform("max").fillna(False).astype(bool)
    return df


# ============================================================================
# 6. 算钱
# ============================================================================
def money(df, cfg):
    df = df.copy()
    FX = P(cfg, "USD_CNY", 6.65)
    MULT = P(cfg, "DEFAULT_SERVICE_FEE", 1.15)

    # 服务费：美国列 > 普通列 > 国家毛利 > 默认
    is_us = df["国家"].astype(str).str.startswith("美国")
    svc = (df["服务费_美国"].where(is_us & df["服务费_美国"].notna(), df["服务费"])
           .fillna(df["国家"].map(cfg["margins"]))
           .fillna(MULT))
    df["服务费"] = svc

    # 附加费 / 欧盟税
    df["附加费"] = df["国家"].map(
        lambda c: (cfg["country"].get(c) or {}).get("surcharge", 0.0)).fillna(0.0)
    eu_tax = round(P(cfg, "EU_TAX_EUR", 3) / P(cfg, "EUR_USD", 0.9), 2)
    df["欧盟税"] = df["国家"].map(
        lambda c: eu_tax if (cfg["country"].get(c) or {}).get("is_eu") else 0.0).fillna(0.0)

    # 国内运费：N 件一价
    fixed_flag = df["账单固定金额"].notna()
    key = ["交易编号", "配置组"]        # 同一配置行的别名合并计费单位
    qty_of_sku = df.groupby(key)["商品数量"].transform("sum")
    first_of_sku = df.groupby(key).cumcount() == 0
    packs = np.ceil(qty_of_sku / df["每计费件数"]).astype(int)
    df["国内运费份数"] = packs.where(first_of_sku, 0)
    df["国内运费合计"] = df["国内运费"].fillna(0) * df["国内运费份数"]
    df.loc[df["非实物"], ["国内运费份数", "国内运费合计"]] = 0

    # 运费：逐行；挂号费整单一次，落在第一个非固定金额商品行
    freight_row = pd.Series(False, index=df.index)
    for _txn, g in df.groupby("交易编号", sort=False):
        cand = g.index[~fixed_flag.loc[g.index]]
        if len(cand):
            freight_row.loc[cand[0]] = True
    df["运费行"] = freight_row

    df["行运费"] = (df["公斤价"] * df["行重量"]).fillna(0)
    df.loc[fixed_flag, "行运费"] = 0.0
    df["挂号费合计"] = df["挂号费"].fillna(0).where(df["运费行"], 0)
    df["运费"] = df["行运费"] + df["挂号费合计"]

    # 人民币
    df["人民币"] = (df["运费"] + df["商品数量"] * df["采购价"]
                    + df["国内运费合计"]
                    + df["附加费"].where(df["首行"], 0)) * df["服务费"]
    df["美元"] = df["人民币"] / FX

    # 固定金额：只锁该行（= 每计费单位价格 × 份数）
    fixed_amt = pd.to_numeric(df["账单固定金额"], errors="coerce")
    df["固定金额合计"] = (fixed_amt * df["国内运费份数"]).round(6)
    df["人民币"] = df["人民币"].where(fixed_amt.isna(), df["固定金额合计"] * FX)
    df["美元"] = df["人民币"] / FX

    # 账单
    rnd = int(P(cfg, "BILL_ROUND", 2))
    df["账单金额"] = (df.groupby("交易编号", sort=False)["美元"].transform("sum")
                      + df["欧盟税"]).round(rnd)
    return df


# ============================================================================
# 7. 异常
# ============================================================================
def build_issues(df, cfg):
    out = []
    for txn, g in df.groupby("交易编号", sort=False):
        h = g.iloc[0]
        r0 = int(h["源表行号"])
        ctry = h["国家"]
        if h["缺运价"]:
            out.append(("阻断", txn, h["订单编号"], r0, "MISSING_FREIGHT",
                        f"【无法查询·待补录运价】「{ctry}」无 {h['订单重量']}kg 的运价"))
        if ctry in cfg["country"] and cfg["country"][ctry]["is_eu"]:
            out.append(("警告", txn, h["订单编号"], r0, "TAX_DATE_FALLBACK",
                        f"欧盟国家「{ctry}」按批次日期处理；欧盟税固定 "
                        f"{P(cfg, 'EU_TAX_EUR', 3)}€/单"))
        if ctry == "美国（偏远）":
            out.append(("警告", txn, h["订单编号"], r0, "US_REMOTE_CHANNEL",
                        f"渠道含「岛屿」，国家改写为「{ctry}」"))
        if ctry == "美国" and not str(h.get("内部单号", "") or "").strip():
            out.append(("警告", txn, h["订单编号"], r0, "US_REMOTE_NO_INTERNAL",
                        "美国订单无内部单号，疑似偏远漏标"))
        unconf = g[g["前缀"].isna() & ~g["非实物"]]
        for i, x in unconf.iterrows():
            out.append(("警告", txn, h["订单编号"], int(x["源表行号"]),
                        "SKU_FALLBACK_DEFAULT",
                        f"SKU「{x['SKU']}」未在产品配置中，使用兜底价"))
    return out


# ============================================================================
# 8. 汇总
# ============================================================================
def summarize(df, cfg):
    FX = P(cfg, "USD_CNY", 6.65)
    rows = []
    for txn, g in df.groupby("交易编号", sort=False):
        h = g.iloc[0]
        if h["缺运价"]:
            continue
        qty = int(g.loc[~g["非实物"], "商品数量"].sum())
        rows.append([h["国家"], txn, qty, h["订单重量"], g["运费"].sum(),
                     g["国内运费合计"].sum(),
                     (g["商品数量"] * g["采购价"]).sum(),
                     h["附加费"], round(h["欧盟税"] * FX, 2), h["账单金额"]])
    t = pd.DataFrame(rows, columns=["国家", "交易编号", "件数", "总重量", "国际运费",
                                    "国内运费", "采购价", "附加费", "欧盟税", "账单合计"])
    agg = t.groupby("国家", sort=False).agg(
        订单数=("交易编号", "nunique"), 件数=("件数", "sum"), 总重量=("总重量", "sum"),
        国际运费=("国际运费", "sum"), 国内运费=("国内运费", "sum"),
        采购价=("采购价", "sum"), 附加费=("附加费", "sum"),
        欧盟税=("欧盟税", "sum"), 账单合计=("账单合计", "sum")).round(2).reset_index()
    total = ["合计"] + [round(agg[c].sum(), 2) for c in agg.columns if c != "国家"]
    return agg, total


# ============================================================================
# 入口：算一遍并打印结果
# ============================================================================
def run(order_path, data_dir=None):
    cfg = load_cfg(data_dir)
    df = read_orders(order_path)
    df = fix_country(df, cfg)
    df = match_sku(df, cfg)
    df = match_rate(df, cfg)
    df = money(df, cfg)
    issues = build_issues(df, cfg)
    agg, total = summarize(df, cfg)
    return df, issues, agg, total


if __name__ == "__main__":
    import sys
    order = sys.argv[1] if len(sys.argv) > 1 else \
        "/Users/zuo/n8n/files/data/workflow003/input/订单导出数据_20261113120442.xlsx"
    df, issues, agg, total = run(order)
    print(f"明细 {len(df)} 行 / {df['交易编号'].nunique()} 单，异常 {len(issues)} 条")
    print("汇总：", total)
