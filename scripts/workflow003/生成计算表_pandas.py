# -*- coding: utf-8 -*-
"""
60103 计算表 · pandas 版
pandas 负责「读 + 清洗 + 配对 + 汇总」，openpyxl 负责「按模板写」。

写表那一步直接复用 生成计算表.py 里的 write()，一行没改。
运行：python 生成计算表_pandas.py
"""
import datetime
import os
import re

import pandas as pd

# from 生成计算表 import EU, EU_TAX, FX, MULT, SUR, write   # 复用口径常量与写表函数

S = "./data"
ORDER = "./data/订单导出数据20261104182737-623523-652214512317620224.xlsx"
FREIGHT = "./data/60103物流运费-11.2更.xlsx"
PRODUCT = "./data/商品配置-11.2更.xlsx"

NOW = datetime.datetime.now()
OUT = f"输出/60103计算表_pandas_{NOW:%Y%m%d_%H%M%S}.xlsx"

FX = 6.65                                  # 美元汇率：AC = AA / FX
MULT = 1.15                                # 服务费：AA 先 ×1.15
EUR, EUR2USD, EU_TAX = 3, 0.9, 3.33        # 欧盟税三列 T / U / V
EU = {"德国", "法国", "意大利", "西班牙", "荷兰", "比利时", "卢森堡", "奥地利",
      "葡萄牙", "爱尔兰", "丹麦", "瑞典", "芬兰", "希腊", "波兰", "捷克", "匈牙利",
      "罗马尼亚", "保加利亚", "克罗地亚", "斯洛文尼亚", "斯洛伐克", "爱沙尼亚",
      "拉脱维亚", "立陶宛", "马耳他", "塞浦路斯"}       # 英国已脱欧
SUR = {"墨西哥": 2.9, "智利": 7, "卡塔尔": 41, "沙特": 40, "阿联酋": 2}
FILL = {"blocked": "FFFFC7CE", "eu": "FFDDEBF7", "remote": "FFE4DFEC", "surcharge": "FFFCE4D6"}


# 读取商品配置文件 -> 读取物流运费表，清晰出国家、重量上下界、公斤价、挂号费等有效字段 -> 


# ============================ 1. 读成 DataFrame ============================
def read_products():
    """
    商品配置 → DataFrame(前缀, 采购价, 偏远采购价, 国内运费, 偏远国内运费, 重量)。

    输入：商品配置.xlsx 的「商品信息」表
    注意：表头是两层合并的，用 header=None + skiprows=3 跳过，列名自己起；
          同一前缀可能有多行，必须去重，否则后面 merge 会让订单行翻倍
    """
    df = pd.read_excel(PRODUCT, sheet_name="商品信息", header=None,
                       skiprows=3, usecols="C:H")
    df.columns = ["前缀", "采购价", "偏远采购价", "国内运费", "偏远国内运费", "重量"]
    # 同一前缀可能有多行（如 CHERRY-DOT），只留第一行，否则 merge 会把订单行翻倍
    return df.dropna(subset=["前缀"]).drop_duplicates(subset=["前缀"], keep="first")


def weight_seg(text):
    """
    重量段文字 → (下界, 上界)；没有重量段返回 None。

    输入：运价表「重量段 KG」整列（会被 .map(weight_seg) 映射）
    输出：元组，或 None（None 的行后面用 dropna 丢掉）
    规则：与 openpyxl 版 seg() 完全一致，区间按「下界开、上界闭」使用
    """
    t = re.sub(r"(?i)kg|\s", "", str(text or ""))
    if "＜W≤" in t:
        a, b = t.split("＜W≤")
    elif "-" in t:
        a, b = t.split("-")
    else:
        return None
    return float(a), float(b)


def read_freight():
    """
    运价表 → DataFrame(国家, 下界, 上界, 公斤价, 挂号费)。

    流程：整表读成字符串 → 国家 ffill → 价格 to_numeric(errors='coerce') →
          拆重量段 → dropna → 按 国家+下界 排序
    注意：'未查询'、空价格会变成 NaN，被 dropna 一次清掉；
          没有重量段的备注行（英国首行 30/16）同样在 dropna 里被清掉
    """
    df = pd.read_excel(FREIGHT, sheet_name=0, dtype=str)
    df["国家"] = df["60103国家报价"].ffill()               # 国家只在首行
    df = df[["国家", "重量段 KG", "公斤价", "挂号费"]].copy()
    for c in ("公斤价", "挂号费"):
        df[c] = pd.to_numeric(df[c], errors="coerce")       # '未查询'/空 → NaN
    df["区间"] = df["重量段 KG"].map(weight_seg)
    df = df.dropna(subset=["国家", "公斤价", "挂号费", "区间"])
    df["下界"] = df["区间"].map(lambda x: x[0])
    df["上界"] = df["区间"].map(lambda x: x[1])
    return df[["国家", "下界", "上界", "公斤价", "挂号费"]].sort_values(["国家", "下界"])


def read_orders():
    """
    订单导出 → DataFrame，一行 = 一件商品。

    流程：dtype=str 整表字符串 → 删空行 → 订单级字段 ffill → 件数转 int →
          记源表行号 → 用 ne(shift()) 标出每单首行
    注意：① 必须 dtype=str，否则 3648 变整数、订单号可能丢精度、邮编丢前导零；
          ② ffill 之后首单自己的空值仍是 NaN，要再 fillna("")，否则后面 .strip() 报错
    """
    df = pd.read_excel(ORDER, dtype=str)                    # dtype=str 保住 3648、7280516530315
    df = df[df["SKU"].notna()].copy()                       # 过滤空行（index 保留，行号不丢）
    for c in ("交易编号", "订单编号", "国家(中)", "物流公司",
              "物流渠道", "内部单号", "邮政编码"):
        df[c] = df[c].ffill().fillna("")                    # 订单级字段只在首行（缺失补空串）
    df["件数"] = df["商品数量"].astype(float).astype(int)
    df["源表行号"] = df.index + 1
    df["首行"] = df["交易编号"].ne(df["交易编号"].shift())   # 每单第一行 = True
    return df


# ============================ 2. 用 DataFrame 算 ============================
def au_zone(post):
    """
    澳洲邮编 → 1 / 2 / 3 区（★占位规则，同 openpyxl 版）。

    输入：订单邮政编码字符串
    输出：区号整数
    注意：正式使用前必须换成云途官方分区表
    """
    pc = int(re.sub(r"\D", "", str(post))[:4] or 0)
    return 1 if 6000 <= pc <= 6999 else 3 if 4000 <= pc <= 4999 else 2


def fix_country(df):
    """
    国家改写：渠道含「岛屿」的美国单 → 美国（偏远）；澳洲 → 澳大利亚（X区）。

    输入/输出：同一个 DataFrame（原地改「国家」列，并新增「偏远」布尔列）
    规则：改写后美国（偏远）单在 match_sku 里会改用偏远那套价格
    注意：str.contains 必须带 na=False，否则 NaN 会让整列掩码出问题
    """
    df["国家"] = df["国家(中)"]
    remote = (df["国家"] == "美国") & df["物流渠道"].str.contains("岛屿", na=False)
    df.loc[remote, "国家"] = "美国（偏远）"
    au = df["国家"] == "澳大利亚"
    df.loc[au, "国家"] = df.loc[au, "邮政编码"].map(lambda p: f"澳大利亚（{au_zone(p)}区）")
    df["偏远"] = df["国家"] == "美国（偏远）"
    return df


def match_sku(df, products):
    """
    SKU → 配置前缀，然后把价格/重量 merge 进来并落成具体数值列。

    输出：新增 前缀 / 采购价 / 国内运费 / 单件重量 等列
    规则：最长前缀匹配；美国（偏远）单取“偏远”两列，其余取普通列；匹配不到走兜底
    注意：merge 前后断言 len(df) 不变 —— 配置里前缀重复会让订单行静默翻倍（踩过）
    """
    prefixes = sorted(products["前缀"], key=len, reverse=True)

    def pick(sku):
        """
        取最长命中的配置前缀，匹配不到返回 None
        """
        for p in prefixes:
            if sku == p or sku.startswith(p + "-"):
                return p
        return None

    df["前缀"] = df["SKU"].map(pick)
    n = len(df)
    df = df.merge(products, on="前缀", how="left")
    assert len(df) == n, f"merge 后行数从 {n} 变成 {len(df)}，检查商品配置是否有重复前缀"
    df["采购价"] = df["偏远采购价"].where(df["偏远"] & df["偏远采购价"].notna(), df["采购价"])
    df["国内运费"] = df["偏远国内运费"].where(df["偏远"] & df["偏远国内运费"].notna(), df["国内运费"])
    df["采购价"] = df["采购价"].fillna(15)                   # 兜底
    df["国内运费"] = df["国内运费"].fillna(4)
    df["单件重量"] = df["重量"].fillna(0.17)
    return df


def price_of(rates, country, w):
    """
    按重量取 (公斤价, 挂号费)，逻辑同 openpyxl 版 rate_of()。

    输入：read_freight() 的 DataFrame、国家、整单重量
    输出：(公斤价, 挂号费)；缺价返回 (None, None)
    """
    s = rates[rates["国家"] == country]
    if s.empty:
        return None, None
    hit = s[(s["下界"] < w) & (w <= s["上界"])]
    if len(hit):
        return hit.iloc[0]["公斤价"], hit.iloc[0]["挂号费"]
    if w < s["下界"].min():
        return s.iloc[0]["公斤价"], s.iloc[0]["挂号费"]
    return None, None


def calc(df, rates):
    """
    整表一次算完金额（向量化）。

    流程：行重量 → 整单重量 transform("sum") → 每单查一次运价再 merge 广播 →
          附加费 / 欧盟税 / 运费 / 人民币 / 美元 / 账单价格
    规则：运费用 .where(首行, 0) 只留在订单首行；
          账单价格 = 整单美元之和 + 欧盟税，用 transform 广播回每一行
    """
    df["行重量"] = df["件数"] * df["单件重量"]
    df["订单重量"] = df.groupby("交易编号", sort=False)["行重量"].transform("sum").round(3)

    # 每单算一次运价，再广播回每一行
    heads = df[df["首行"]].copy()
    pairs = heads.apply(lambda r: price_of(rates, r["国家"], r["订单重量"]), axis=1)
    heads["公斤价"] = [p[0] for p in pairs]
    heads["挂号费"] = [p[1] for p in pairs]
    n = len(df)
    df = df.merge(heads[["交易编号", "公斤价", "挂号费"]], on="交易编号", how="left")
    assert len(df) == n, f"merge 后行数从 {n} 变成 {len(df)}，检查是否有重复交易编号"

    df["附加费"] = df["国家"].map(SUR).fillna(0)
    df["欧盟税"] = df["国家"].isin(EU).astype(float) * EU_TAX
    df["运费"] = (df["公斤价"] * df["订单重量"] + df["挂号费"]).fillna(0).where(df["首行"], 0)
    df["人民币"] = (df["运费"] + df["件数"] * (df["采购价"] + df["国内运费"])
                    + df["附加费"].where(df["首行"], 0)) * MULT
    df["美元"] = df["人民币"] / FX
    df["帐单金额"] = (df.groupby("交易编号", sort=False)["美元"].transform("sum")
                      + df["欧盟税"]).round(2)
    return df


def to_result(df):
    """
    DataFrame → write() 要的 dict 结构（一个订单一个 dict）。

    输出：(result, issues)，格式与 openpyxl 版 calc() 的返回完全一致
    注意：这是全流程唯一“退回 Python 循环”的地方 —— 因为 write() 是逐格写的，
          DataFrame 必须在这里拆成每行元组；pandas 的收益到这里就结束了
    """
    result, issues = [], []
    for _, g in df.groupby("交易编号", sort=False):
        h = g.iloc[0]
        blocked = pd.isna(h["公斤价"])
        notes = []
        if h["国家"] == "美国（偏远）":
            notes.append("美国（偏远）专线")
        if h["欧盟税"]:
            notes.append(f"欧盟国家，收欧盟税 {EU_TAX}$")
        if h["附加费"]:
            notes.append(f"含附加费 {h['国家']} +{h['附加费']:.1f}￥")
        rows = [(int(x["件数"]), x["SKU"], x["采购价"], x["国内运费"], x["单件重量"],
                 x["人民币"], x["美元"]) for _, x in g.iterrows()]
        result.append(dict(
            o={"txn": h["交易编号"], "no": h["订单编号"], "logi": h["物流公司"],
               "chan": h["物流渠道"], "inner": h["内部单号"] or "", "src": h["源表行号"]},
            country=h["国家"], weight=h["订单重量"], price=h["公斤价"], fee=h["挂号费"],
            blocked=blocked, sur=h["附加费"], tax=h["欧盟税"], rows=rows,
            total=float(g["美元"].sum()), bill=float(h["帐单金额"]), notes=notes))

        if blocked:
            issues.append(("阻断", h["交易编号"], h["订单编号"], h["源表行号"],
                           f"「{h['国家']}」缺 {h['订单重量']}kg 运价"))
        if h["国家"] == "美国" and not (h["内部单号"] or "").strip():
            issues.append(("警告", h["交易编号"], h["订单编号"], h["源表行号"],
                           "美国单无内部单号，疑似偏远漏处理"))
        for _, x in g[g["前缀"].isna()].iterrows():
            issues.append(("警告", h["交易编号"], h["订单编号"], x["源表行号"],
                           f"SKU「{x['SKU']}」未在产品配置中"))
    return result, issues





if __name__ == "__main__":
    products = read_products()
    print("=====================第一步=====================")
    print(products)
    rates = read_freight()
    print("=====================第二步=====================")
    print(rates)

    orders = fix_country(read_orders())                # ① 读 + 国家改写
    print("=====================第三步=====================")
    print(orders)
    detail = calc(match_sku(orders, products), rates)  # ② 配价格 → ③ 算金额
    print("=====================第四步=====================")
    print(detail)
    result, issues = to_result(detail)
    path = write(result, issues, out=OUT, now=NOW)

    print(f"✔ {path}")
    print(f"  {len(result)} 单 / {len(detail)} 行 / 阻断 {sum(g['blocked'] for g in result)} 单")
    for level, txn, no, _src, msg in issues:
        print(f"  [{level}] {txn} {no} {msg}")
