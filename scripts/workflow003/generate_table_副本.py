import datetime
import os
import io
import re
import shutil
import zipfile
from copy import copy

import openpyxl
import pandas as pd
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image
import urllib.request
from openpyxl.drawing.image import Image as XLImage




# ============================================================================
# 业务参数 —— 改口径只改这一段
# ============================================================================

# ---------- 1. 汇率 ----------
FX = 6.65                        # 人民币 / 美元：AC = AA / FX、W = V * FX

# ---------- 2. 服务费（毛利）----------
SERVICE_COL = "毛利"             # 运价表里存服务费率的列名
MULT = 1.15                      # 默认服务费：该国毛利列没填时用这个

# ---------- 3. 欧盟税（每单一次，不乘服务费）----------
EUR = 3.0                        # 欧盟国家每单多收 3 欧元
EUR2USD = 0.9                    # 欧元 → 美元汇率（会变）
EU_TAX = round(EUR / EUR2USD, 2) # 折成美元 = 3.33（V 列真正参与计算的是它）

# 欧盟名单：英国已脱欧；挪威 / 冰岛 / 瑞士不是欧盟
EU = {"德国", "法国", "意大利", "西班牙", "荷兰", "比利时", "卢森堡", "奥地利",
      "葡萄牙", "爱尔兰", "丹麦", "瑞典", "芬兰", "希腊", "波兰", "捷克", "匈牙利",
      "罗马尼亚", "保加利亚", "克罗地亚", "斯洛文尼亚", "斯洛伐克", "爱沙尼亚",
      "拉脱维亚", "立陶宛", "马耳他", "塞浦路斯"}

# ---------- 4. 附加费（人民币 / 单，加在订单首行）----------
# 0 表示"规则有但当前不收"；要收就改数字，不用动代码
SUR = {
    "墨西哥": 2.9,
    "智利": 7.0,
    "卡塔尔": 41.0,
    "沙特": 40.0,
    "阿联酋": 0.0,        # 规则是 2 元，但「现在还没收」；开始收改成 2.0
    "罗马尼亚": 42.0,     # 处理费
}

# ---------- 5. 行底色（不参与计算）----------
FILL = {"blocked": "FFFFC7CE", "eu": "FFDDEBF7", "remote": "FFE4DFEC", "surcharge": "FFFCE4D6"}

QTY = "商品数量"                  # 你的件数列名
WT = "重量"                       # 你的单件重量列名


# ---- 商品图 ----
# "url"      = 按「原始图片」列的链接下载、缩小后插到 G 列（Excel/WPS 都能看）
# "dispimg"  = 用模板里的 WPS 单元格图片公式（只有 WPS 能显示）
IMAGE_MODE = "url"
IMG_URL_COL = "原始图片"
IMG_MAX = 60                      # 图片最长边（像素），另一边按原比例算
_thumb_cache = {}


# 注意点：
# 1、物流运费表拆分重量字段时，必须要是-或者是< ≤
# 2、物流运费表和商品配置表字段顺序必须保持一致，更换后可能导致错乱



pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

BASE = os.environ.get("BILL_BASE") or os.path.dirname(os.path.abspath(__file__))
S = os.path.join(BASE, "data")
TEMPLATE = os.path.join(BASE, "data", "60103计算表_20261104_183110.xlsx")     # 结果模板（以它为底）
OUT_DIR = os.environ.get("BILL_OUT") or os.path.join(BASE, "输出")

# 导出的订单表
ORDER = os.path.join(S, "订单导出数据_20261107040803.xlsx")
# 物流运费表
FREIGHT = os.path.join(S, "60103物流运费-11.6更(1).xlsx")
# 商品配置表
PRODUCT = os.path.join(S, "商品配置-11.6更.xlsx")
# 结果模板表
TEMPLATE = os.path.join(S, "计算结果模板表.xlsx")
# 澳大利亚分区表
AU_ZONE_FILE = os.path.join(S, "澳大利亚分区.xlsx")

# 读取商品配置表并保留需要的数据列
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

# 读取物流运费表保留相对应的字段和区间
def read_freight():
    """
    运价表 → DataFrame(国家, 下界, 上界, 公斤价, 挂号费)。

    流程：整表读成字符串 → 国家 ffill → 价格 to_numeric(errors='coerce') →
          拆重量段 → dropna → 按 国家+下界 排序
    注意：'未查询'、空价格会变成 NaN，被 dropna 一次清掉；
          没有重量段的备注行（英国首行 30/16）同样在 dropna 里被清掉
    """
    df = pd.read_excel(FREIGHT, sheet_name=0, dtype=str)
    df["国家"] = df["60103国家报价"].ffill()

    # ① 选列时把「毛利」带上（列不存在就跳过，兼容旧运费表）
    cols = ["国家", "重量段 KG", "公斤价", "挂号费"]
    if SERVICE_COL in df.columns:
        cols.append(SERVICE_COL)
    df = df[cols].copy()

    # ② 价格类列统一转数字
    for c in ("公斤价", "挂号费"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    if SERVICE_COL in df.columns:
        df[SERVICE_COL] = pd.to_numeric(df[SERVICE_COL], errors="coerce")   # '1.15'/空 → 数字/NaN

    df["区间"] = df["重量段 KG"].map(weight_seg)
    df = df.dropna(subset=["国家", "公斤价", "挂号费", "区间"])      # ③ 注意：这里不放毛利

    df["下界"] = df["区间"].map(lambda x: x[0])
    df["上界"] = df["区间"].map(lambda x: x[1])

    keep = ["国家", "下界", "上界", "公斤价", "挂号费"]
    if SERVICE_COL in df.columns:
        keep.append(SERVICE_COL)                                    # ④ 返回时保留
    return df[keep].sort_values(["国家", "下界"])

# 读取导出的订单数据并标记首行，在后续写入结果表中需要判定
def read_orders():
    """
    订单导出 → DataFrame，一行 = 一件商品。

    流程：dtype=str 整表字符串 → 删空行 → 订单级字段 ffill → 件数转 int →
          记源表行号 → 用 groupby+cumcount 标出每单首行
    注意：① 必须 dtype=str，否则 3648 变整数、订单号可能丢精度、邮编丢前导零；
          ② ffill 之后首单自己的空值仍是 NaN，要再 fillna("")，否则后面 .strip() 报错
    """
    df = pd.read_excel(ORDER, dtype=str)                    # dtype=str 保住 3648、7280516530315

    df = df[df["SKU"].notna()].copy()

    # 将缺失的数据不齐，将nan值都换为空字符串
    for c in ("交易编号", "订单编号", "国家中文", "物流公司", "物流渠道", "内部单号", "邮政编码"):
        df[c] = df[c].ffill().fillna("")

    # 先分组，然后对组内的数据排序（0，1，2）,最后将为0的标为首行
    df["首行"] = df.groupby("交易编号", sort=False).cumcount() == 0

    # 首行必须正好等于订单数；同一单的行必须连续
    assert df["首行"].sum() == df["交易编号"].nunique(), "首行数与订单数不一致"
    assert (df["交易编号"] != df["交易编号"].shift()).sum() == df["交易编号"].nunique(), \
        "同一订单的行不连续，需要先按交易编号稳定排序"
    return df

# 修改订单数据中的国家，分类出偏远等类型
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

# 读取澳大利亚的分区表

_zones = None

def au_zones():
    """所有分区块合并去重，同一邮编取第一个 → {'6175': '1区', ...}"""
    global _zones
    if _zones is not None:
        return _zones
    raw = pd.read_excel(AU_ZONE_FILE, header=None, dtype=str)
    hdr = next(r for r in range(30) if (raw.iloc[r] == "邮编").any())
    blocks = []
    for c in range(raw.shape[1]):
        if str(raw.iat[hdr, c]).strip() != "邮编":
            continue
        blk = raw.iloc[hdr + 1:, [c, c + 1]].copy()
        blk.columns = ["邮编", "分区"]
        blocks.append(blk)
    df = pd.concat(blocks, ignore_index=True)
    df["邮编"] = df["邮编"].str.strip().str.split(".").str[0].str.zfill(4)
    df["分区"] = df["分区"].str.strip()
    df = df[df["邮编"].str.isdigit() & df["分区"].notna() & df["分区"].ne("")]
    df = df.drop_duplicates("邮编", keep="first")
    _zones = dict(zip(df["邮编"], df["分区"]))
    return _zones


def au_zone(post):
    """澳洲邮编 → 1 / 2 / 3；查不到 → None"""
    pc = re.sub(r"\D", "", str(post or ""))
    z = au_zones().get(pc.zfill(4)[-4:]) if pc else None
    return int(z[0]) if z else None

def fix_country(df):
    """
    改写订单的「国家」：美国偏远 + 澳大利亚分区，一次做完。

      ① 美国 且 物流渠道含「岛屿」  → 美国（偏远）     偏远=True，后续用偏远那套价
      ② 澳大利亚 且 邮编在分区表里  → 澳大利亚（X区）  X = 1 / 2 / 3
      ③ 澳大利亚 但邮编查不到      → 保持「澳大利亚」 即"没区号" → 后续阻断，绝不猜区
      ④ 其余                      → 原名不动

    输出：只改「国家」一列 + 新增「偏远」布尔列（选成本用），不再多开分区列
    """
    df = df.rename(columns={"国家中文": "国家"})
    df["国家"] = df["国家"].fillna("")

    # ① 美国 + 岛屿专线 → 美国（偏远）
    remote = df["国家"].eq("美国") & df["物流渠道"].str.contains("岛屿", na=False)
    df.loc[remote, "国家"] = "美国（偏远）"
    df["偏远"] = df["国家"].eq("美国（偏远）")

    # ② 澳大利亚 → 澳大利亚（X区）；查不到就保持「澳大利亚」
    au = df["国家"].eq("澳大利亚")
    if au.any():
        zone = df.loc[au, "邮政编码"].map(au_zone)        # 临时 Series，不落到 df 上
        df.loc[au, "国家"] = zone.map(
            lambda z: f"澳大利亚（{int(z)}区）" if pd.notna(z) else "澳大利亚")
    return df

def match_sku(df, prod):
    """SKU → 前缀 → 采购价/国内运费/单件重量（含偏远选价与兜底）"""
    # df = norm_cols(df)
    prefixes = sorted(prod["前缀"], key=len, reverse=True)
    df["前缀"] = df["SKU"].map(lambda s: next(
        (p for p in prefixes if s == p or s.startswith(p + "-")), None))

    n = len(df)
    df = df.merge(prod, on="前缀", how="left")
    assert len(df) == n, f"merge 后行数从 {n} 变成 {len(df)}，检查配置是否重复前缀"

    df["采购价"] = df["偏远采购价"].where(df["偏远"] & df["偏远采购价"].notna(), df["采购价"])
    df["国内运费"] = df["偏远国内运费"].where(df["偏远"] & df["偏远国内运费"].notna(), df["国内运费"])
    df["价格缺失"] = df["前缀"].notna() & df["采购价"].isna()
    df["采购价"] = df["采购价"].fillna(15)
    df["国内运费"] = df["国内运费"].fillna(4)
    df["重量"] = df["重量"].fillna(0.17) 
    df["单件重量"] = df["重量"].fillna(0.17)

    # 非实物行：SKU 以 SHIP 开头且没有商品图 → 不计重量、不计价格
    no_img = df["原始图片"].fillna("").astype(str).str.strip().eq("")
    df["非实物"] = df["SKU"].astype(str).str.upper().str.startswith("SHIP") & no_img
    df.loc[df["非实物"], ["采购价", "国内运费", "重量", "单件重量"]] = 0.0
    df.loc[df["非实物"], "价格缺失"] = False

    return df

def match_rate(df, rates):
    """
    把运价写回订单明细。

    ① 行重量   = 件数 × 单件重量            ← 件数就是「商品数量」
    ② 整单重量 = 按交易编号求和，transform 广播回每行（行数不变）
    ③ 每单查一次运价（区间匹配）
    ④ merge 回明细，标记「缺运价」

    输入：df 要有 件数 / 单件重量 / 国家 / 交易编号 / 首行
    输出：新增 行重量 / 订单重量 / 公斤价 / 挂号费 / 缺运价
    """
    df["行重量"] = df["商品数量"] * df["重量"]
    df["订单重量"] = df.groupby("交易编号", sort=False)["行重量"].transform("sum").round(3)

    heads = df.loc[df["首行"], ["交易编号", "国家", "订单重量"]].drop_duplicates("交易编号")
    pairs = heads.apply(lambda r: price_of(rates, r["国家"], r["订单重量"]), axis=1)
    heads["公斤价"] = [p[0] for p in pairs]
    heads["挂号费"] = [p[1] for p in pairs]
    heads["缺运价"] = heads["公斤价"].isna()

    n = len(df)
    df = df.merge(heads[["交易编号", "公斤价", "挂号费", "缺运价"]], on="交易编号", how="left")
    assert len(df) == n, f"merge 后行数从 {n} 变成 {len(df)}，检查是否有重复交易编号"
    return df

def add_money(df, svc_map):
    """
    在已融合运价的明细上算钱。

    输入：df 需有 件数 / 采购价 / 国内运费 / 国家 / 首行 / 交易编号 /
          订单重量 / 公斤价 / 挂号费
    输出：新增 服务费 / 附加费 / 欧盟税 / 运费 / 人民币 / 美元 / 账单金额
    """
    df["服务费"] = df["国家"].map(svc_map).fillna(MULT)
    df["附加费"] = df["国家"].map(SUR).fillna(0)
    df["欧盟税"] = df["国家"].isin(EU).astype(float) * EU_TAX

    # 运费：整单一次，只留在首行
    df["运费"] = (df["公斤价"] * df["订单重量"] + df["挂号费"]).fillna(0).where(df["首行"], 0)

    # 人民币总价：运费 + 成本 + 附加费，再乘服务费
    df["人民币"] = (df["运费"] + df["件数"] * (df["采购价"] + df["国内运费"])
                    + df["附加费"].where(df["首行"], 0)) * df["服务费"]

    # 折美元
    df["美元"] = df["人民币"] / FX

    # 账单金额：整单美元之和 + 欧盟税（每单一次），保留 2 位
    df["账单金额"] = (df.groupby("交易编号", sort=False)["美元"].transform("sum")
                      + df["欧盟税"]).round(2)
    return df

def add_money(df, svc_map):
    """
    在已融合运价的明细上算钱。

    输入：df 需有 件数 / 采购价 / 国内运费 / 国家 / 首行 / 交易编号 /
          订单重量 / 公斤价 / 挂号费
    输出：新增 服务费 / 附加费 / 欧盟税 / 运费 / 人民币 / 美元 / 账单金额
    """

    # print(df.dtypes[df.dtypes == "object"])

    df = df.copy()

    for c in ("采购价", "国内运费", "偏远采购价", "偏远国内运费"):
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    df["服务费"] = df["国家"].map(svc_map).fillna(MULT)
    df["附加费"] = df["国家"].map(SUR).fillna(0)
    df["欧盟税"] = df["国家"].isin(EU).astype(float) * EU_TAX

    # 运费：整单一次，只留在首行
    df["运费"] = (df["公斤价"] * df["订单重量"] + df["挂号费"]).fillna(0).where(df["首行"], 0)

    # 人民币总价：运费 + 成本 + 附加费，再乘服务费
    df["人民币"] = (df["运费"] + df["商品数量"] * (df["采购价"] + df["国内运费"])
                    + df["附加费"].where(df["首行"], 0)) * df["服务费"]

    # 折美元
    df["美元"] = df["人民币"] / FX

    # 账单金额：整单美元之和 + 欧盟税（每单一次），保留 2 位
    df["账单金额"] = (df.groupby("交易编号", sort=False)["美元"].transform("sum")
                      + df["欧盟税"]).round(2)
    return df




# 拆分重量段 KG这个字段，方便比较（在物流运费表中使用）
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

#  传入物流运费表数据，国家， 重量数值来获取匹配的公斤价和挂号费
def price_of(rates, country, w):
    """重量 → (公斤价, 挂号费)：命中段 / 更轻按最低段 / 更重返回 NaN"""
    s = rates[rates["国家"] == country]
    if s.empty:
        return None, None
    hit = s[(s["下界"] < w) & (w <= s["上界"])]               # 下界开、上界闭
    if len(hit):
        return hit.iloc[0]["公斤价"], hit.iloc[0]["挂号费"]
    if w < s["下界"].min():                                  # 比最低段还轻
        return s.iloc[0]["公斤价"], s.iloc[0]["挂号费"]
    return None, None      

# 将没填毛利的国家都用一个毛利
def service_rates(rates):
    """国家 → 服务费率（运价表「毛利」列）；没填的国家由调用处用 MULT 兜底"""
    if SERVICE_COL not in rates.columns:
        return {}
    return rates.groupby("国家")[SERVICE_COL].first().dropna().to_dict()




# ================================================写入结果表====================================================================

# ============================================================================
# 1. 异常清单（6 类全部捕获）
# ============================================================================
# 6 类异常代码（与老计算表一致）
#   MISSING_FREIGHT        阻断  该国家+线路下找不到覆盖整单重量的运价段
#   US_REMOTE_NO_INTERNAL  警告  美国单但内部单号为空（疑似偏远漏标）
#   US_REMOTE_CHANNEL      警告  渠道含「岛屿」→ 国家改写为 美国（偏远）
#   ROUTE_FALLBACK         警告  该国没配线路映射，但运价表里只有一条线路
#   TAX_DATE_FALLBACK      警告  欧盟单但订单表没有「创建日期」列
#   SKU_FALLBACK_DEFAULT   警告  SKU 匹配不到商品配置，用了默认价

DEFAULT_PURCHASE = 15          # SKU 未配置时的默认采购价（写进异常说明用）


def build_issues(df, rates=None, route_map=None, batch_date=None):
    """
    从明细 df 推导全部 6 类异常。

    输入：
        df          你的明细 DataFrame（要有 交易编号/订单编号/国家/SKU/前缀/
                    内部单号/物流渠道/物流公司/首行/缺运价/订单重量）
        rates       运价表 DataFrame（可选；给了才能判断 ROUTE_FALLBACK 和写出线路名）
        route_map   {国家: 线路} 映射表（可选；没配的国家才报 ROUTE_FALLBACK）
        batch_date  批次日期（可选；TAX_DATE_FALLBACK 的说明里用，默认今天）
    输出：
        [(级别, 交易编号, 订单号, 源表行号, 代码, 说明), ...]
        顺序：按订单在 df 里出现的顺序，同单内 阻断 → 线路 → 税 → 偏远 → 无单号 → SKU
    """
    if "源表行号" in df.columns:
        src = df["源表行号"]
    else:                                        # 没有行号列 → 用行序兜底
        src = pd.Series(range(1, len(df) + 1), index=df.index)

    route_map = route_map or {}
    batch = batch_date or f"{pd.Timestamp.now():%Y-%m-%d}"
    has_date_col = any(c in df.columns for c in ("创建日期", "订单创建时间"))

    # 每个国家的「线路数」和「线路名」，用于 ROUTE_FALLBACK
    route_cnt, route_name = {}, {}
    if rates is not None and "线路" in getattr(rates, "columns", []):
        route_cnt = rates.groupby("国家")["线路"].nunique().to_dict()
        route_name = rates.groupby("国家")["线路"].first().to_dict()

    issues, warned_route = [], set()             # warned_route：线路回退每个国家只报一次
    for txn, g in df.groupby("交易编号", sort=False):
        h = g.iloc[0]
        row0 = int(src.loc[g.index[0]])
        ctry = h["国家"]

        # ① 阻断：缺运价
        if h["缺运价"]:
            issues.append(("阻断", txn, h["订单编号"], row0, "MISSING_FREIGHT",
                           f"【无法查询·待补录运价】「{ctry}」无 {h['订单重量']}kg 的运价"))

        # ② 线路回退（每国只报第一条订单）
        if route_cnt.get(ctry) == 1 and ctry not in route_map and ctry not in warned_route:
            warned_route.add(ctry)
            issues.append(("警告", txn, h["订单编号"], row0, "ROUTE_FALLBACK",
                           f"国家「{ctry}」未配置线路映射，使用该国唯一线路「{route_name.get(ctry, '唯一线路')}」"))

        # ③ 欧盟税：订单表没有创建日期列
        if ctry in EU and not has_date_col:
            issues.append(("警告", txn, h["订单编号"], row0, "TAX_DATE_FALLBACK",
                           f"欧盟国家「{ctry}」订单表无创建日期列，已按批次日期（{batch}）处理；"
                           f"税额不受影响（欧盟税当前固定 {EU_TAX}$/单，「每月25号后加收5.5$」条款未启用）"))

        # ④ 美国偏远改写
        if ctry == "美国（偏远）":
            issues.append(("警告", txn, h["订单编号"], row0, "US_REMOTE_CHANNEL",
                           f"物流公司含「{h.get('物流公司', '')}」且渠道含「岛屿」，国家改写为「{ctry}」"))

        # ⑤ 美国单无内部单号
        if ctry == "美国" and not str(h.get("内部单号", "") or "").strip():
            issues.append(("警告", txn, h["订单编号"], row0, "US_REMOTE_NO_INTERNAL",
                           "美国订单无内部单号，疑似属于美国（偏远）漏处理，请核实"
                           "（当前仅提醒不改写，计费按普通美国价）"))

        # ⑥ SKU 未配置（商品级，行号指向该 SKU 所在行）
        for i, x in g[g["前缀"].isna()].iterrows():
            issues.append(("警告", txn, h["订单编号"], int(src.loc[i]), "SKU_FALLBACK_DEFAULT",
                           f"SKU「{x['SKU']}」未在产品配置中，使用默认价格（采购{DEFAULT_PURCHASE:g}）"))
    return issues


# ============================================================================
# 1.5 商品图：按 URL 下载 → 等比缩小 → 插到 G 列
# ============================================================================
def _thumb(url, size=IMG_MAX):
    """下载图片并缩成小图（同一个 URL 只下一次），返回 PNG bytes"""
    if url in _thumb_cache:
        return _thumb_cache[url]
    with urllib.request.urlopen(url, timeout=20) as resp:
        raw = resp.read()
    im = Image.open(io.BytesIO(raw)).convert("RGB")
    im.thumbnail((size * 2, size * 2))            # 等比缩略（PIL 会保持原比例）
    buf = io.BytesIO()
    im.save(buf, format="PNG", optimize=True)
    _thumb_cache[url] = buf.getvalue()
    return _thumb_cache[url]


def fit_size(w, h, box=IMG_MAX):
    """按原比例缩放：最长边 = box，另一边跟着比例走（横图宽=box，竖图高=box）"""
    scale = box / max(w, h)
    return max(1, round(w * scale)), max(1, round(h * scale))


def insert_images(ws, records, col="G"):
    """records = [(行号, 图片URL)]；下载后锚定到该行 G 列并等比缩小"""
    ok = fail = 0
    for r, url in records:
        if not isinstance(url, str) or not url.startswith("http"):
            continue
        try:
            img = XLImage(io.BytesIO(_thumb(url)))
            img.width, img.height = fit_size(img.width, img.height)   # 保持原比例
            ws.add_image(img, f"{col}{r}")
            ok += 1
        except Exception as e:
            fail += 1
            print(f"  ⚠️ 第 {r} 行图片没插上：{type(e).__name__} {e}")
    print(f"  商品图：成功 {ok} 张，失败 {fail} 张")
    return ok, fail


# ============================================================================
# 2. 行底色分类
# ============================================================================
def cat_of(g):
    """一个订单该刷什么底色：阻断 > 欧盟 > 偏远 > 附加费"""
    h = g.iloc[0]
    if h["缺运价"]:
        return "blocked"
    if h["欧盟税"]:
        return "eu"
    if h["国家"] == "美国（偏远）":
        return "remote"
    if h["附加费"]:
        return "surcharge"
    return ""


# ============================================================================
# 3. 写明细（核心）
# ============================================================================
def write_excel(df, issues, out_path):
    """把明细 df 按模板写成计算表，返回输出路径"""
    if "源表行号" not in df.columns:            # 异常表要行号，没有就用行序兜底
        df = df.assign(源表行号=range(1, len(df) + 1))

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    shutil.copy(TEMPLATE, out_path)             # 以模板为底：样式/列宽/图片都在里面
    wb = openpyxl.load_workbook(out_path)
    ws = wb.worksheets[0]
    ws.title = f"{pd.Timestamp.now():%m.%d}计算表"

    cols = [get_column_letter(i) for i in range(1, 39)]
    style = {c: copy(ws[f"{c}2"]._style) for c in cols}       # 数据行样式
    height = ws.row_dimensions[2].height
    img = {}
    if IMAGE_MODE == "dispimg":                              # 用模板的 SKU→图片公式
        img = {r[9].value: r[6].value
               for r in ws.iter_rows(min_row=2) if r[9].value and r[6].value}
    img_rows = []                                            # [(行号, 图片URL)]

    for rng in list(ws.merged_cells.ranges):                 # 拆旧合并
        ws.unmerge_cells(str(rng))
    ws.delete_rows(2, ws.max_row)                            # 清旧数据，只留表头

    r = 2
    for txn, g in df.groupby("交易编号", sort=False):
        first, last = r, r + len(g) - 1
        for i, (_, row) in enumerate(g.iterrows()):
            if IMAGE_MODE == "url":
                img_rows.append((r, row.get(IMG_URL_COL)))
            v = {"A": "", "G": img.get(row["SKU"], ""), "H": int(row[QTY]),
                 "J": row["SKU"], "K": row["采购价"], "L": row["国内运费"],
                 "M": "万邦普货", "O": "带", "Z": row["服务费"],
                 "AB": FX, "AC": f"=AA{r}/AB{r}"}
            if i == 0:                                        # 订单级字段只写首行
                v.update({
                    "B": row["物流公司"], "C": row["物流渠道"], "D": row["内部单号"],
                    "E": row["订单编号"], "F": row["交易编号"], "I": row["国家"],
                    "P": row["公斤价"] if not row["缺运价"] else "未查询",
                    "Q": row["挂号费"] if not row["缺运价"] else "未查询",
                    "R": "=" + "+".join(f"H{first + k}*{w}"
                                        for k, w in enumerate(g[WT])),
                    "S": f"=N(P{r})*N(R{r})+N(Q{r})",
                    "AA": f"=(N(S{r})+N(K{r})*N(H{r})+N(L{r})*N(H{r})+N(Y{r}))*N(Z{r})",
                    "AE": f"=SUM(AC{first}:AC{last})",
                    "AD": f"=N(V{r})+AE{r}",
                    "AF": f"=ROUND(AD{r},2)"})
                if row["欧盟税"]:
                    v.update({"T": EUR, "U": EUR2USD, "V": row["欧盟税"],
                              "W": f"=V{r}*AB{r}"})
                if row["附加费"]:
                    v["Y"] = row["附加费"]
                if row["缺运价"]:
                    v["AL"] = f"阻断：「{row['国家']}」缺 {row['订单重量']}kg 运价"
            else:
                v["AA"] = f"=(N(K{r})*N(H{r})+N(L{r})*N(H{r}))*N(Z{r})"

            fill = cat_of(g)
            for c in cols:
                cell = ws[f"{c}{r}"]
                cell._style = copy(style[c])
                if c in v:
                    cell.value = v[c]
                cell.fill = (PatternFill("solid", fgColor=FILL[fill]) if fill
                             else PatternFill(fill_type=None))
            ws.row_dimensions[r].height = height
            r += 1
        if last > first:                                      # 国家/重量/账单价格按单合并
            for c in ("I", "R", "AF"):
                ws.merge_cells(f"{c}{first}:{c}{last}")

    # ---- 插商品图（URL 模式）----
    if IMAGE_MODE == "url" and img_rows:
        insert_images(ws, img_rows)

    # ---- 总计行 ----
    ws[f"A{r}"], ws[f"V{r}"] = "总计", f"=SUM(V2:V{r - 1})"
    ws[f"AE{r}"], ws[f"AF{r}"] = f"=SUM(AE2:AE{r - 1})", f"=SUM(AF2:AF{r - 1})"

    # ---- 图例 ----
    legend = ["颜色图例：",
              "阻断：缺运价/缺配置等，金额未计（明细格见「未查询」）",
              "警告：需人工确认（原因见异常表）",
              f"欧盟国家：收欧盟税 {EU_TAX}$/单（{EUR:g}欧元 ÷ {EUR2USD:g}，明细 T/U/V 列）",
              "美国（偏远）：专线改写并套偏远列价格",
              "含附加费国家（" + "/".join(f"{k}{v:g}" for k, v in SUR.items() if v > 0) + "）",
              "金额格「未查询」= 运价表中未查到该运价"]
    for i, text in enumerate(legend):
        ws[f"B{r + 3 + i}"] = text

    # ---- 汇总（阻断单不计）----
    rows = []
    for txn, g in df.groupby("交易编号", sort=False):
        h = g.iloc[0]
        if h["缺运价"]:
            continue
        rows.append([h["国家"], txn, int(g[QTY].sum()), h["订单重量"],
                     g["运费"].sum(),
                     (g[QTY] * g["国内运费"]).sum(),
                     (g[QTY] * g["采购价"]).sum(),
                     h["附加费"], round(h["欧盟税"] * FX, 2), h["账单金额"]])
    t = pd.DataFrame(rows, columns=["国家", "交易编号", "件数", "总重量", "国际运费",
                                    "国内运费", "采购价", "附加费", "欧盟税", "账单合计"])
    agg = t.groupby("国家", sort=False).agg(
        订单数=("交易编号", "nunique"), 件数=("件数", "sum"), 总重量=("总重量", "sum"),
        国际运费=("国际运费", "sum"), 国内运费=("国内运费", "sum"),
        采购价=("采购价", "sum"), 附加费=("附加费", "sum"),
        欧盟税=("欧盟税", "sum"), 账单合计=("账单合计", "sum")).round(2).reset_index()
    summary = [[a["国家"], a["订单数"], a["件数"], a["总重量"], a["国际运费"],
                a["国内运费"], a["采购价"], a["附加费"], a["欧盟税"], a["账单合计"]]
               for _, a in agg.iterrows()]
    summary.append(["合计"] + [round(sum(x[j] for x in summary), 2) for j in range(1, 10)])
    fill_rows(wb["汇总"], summary)

    # ---- 异常 ----
    fill_rows(wb["异常"], [[lv, txn, no, src, code, msg]
                           for lv, txn, no, src, code, msg in
                           sorted(issues, key=lambda x: x[0] != "阻断")], ncol=6)

    # ---- 配置快照 ----
    snap = wb["配置快照"]
    now = pd.Timestamp.now()
    for key, val in {"批次时间": f"{now:%Y-%m-%d %H:%M:%S}",
                     "加载时间": f"{now:%Y-%m-%d}",
                     "美元汇率": str(FX),
                     "欧盟税费($)": f"{EU_TAX}（{EUR:g}欧元 ÷ {EUR2USD:g}）",
                     "服务费率/固定额": f"{MULT}（默认；各国取运价表「毛利」列）",
                     "附加费国家": str({k: v for k, v in SUR.items() if v > 0}),
                     "暂不收附加费": "/".join(k for k, v in SUR.items() if v == 0) or "无",
                     }.items():
        for i in range(2, snap.max_row + 1):
            if snap[f"A{i}"].value == key:
                snap[f"B{i}"] = val

    wb.save(out_path)
    if IMAGE_MODE == "dispimg":              # 用模板的 DISPIMG 才需要补 WPS 图片部件
        keep_wps_images(TEMPLATE, out_path)
    return out_path


# ============================================================================
# 4. 纯值表覆盖写入（汇总 / 异常 用）
# ============================================================================
def fill_rows(ws, rows, ncol=10):
    """从第 2 行起覆盖写入；数据比模板少就删掉多余行，多就套第 2 行样式"""
    old = ws.max_row
    for i, row in enumerate(rows):
        for j, val in enumerate(row, 1):
            c = ws.cell(row=2 + i, column=j, value=val)
            if 2 + i > old:
                c._style = copy(ws.cell(row=2, column=j)._style)
    if old > 1 + len(rows):
        ws.delete_rows(2 + len(rows), old - 1 - len(rows))


# ============================================================================
# 5. 把 WPS 单元格图片补回新文件（仅 IMAGE_MODE="dispimg" 时用）
# ============================================================================
def keep_wps_images(tpl, out):
    """openpyxl 存盘会丢掉 xl/cellimages.xml + xl/media/*，这里从模板塞回去"""
    with zipfile.ZipFile(tpl) as z:
        parts = {n: z.read(n) for n in z.namelist()
                 if n in ("xl/cellimages.xml", "xl/_rels/cellimages.xml.rels")
                 or n.startswith("xl/media/")}
    with zipfile.ZipFile(out) as z:
        ct = z.read("[Content_Types].xml").decode()
        rels = z.read("xl/_rels/workbook.xml.rels").decode()

    ct = ct.replace('<Default Extension="rels"',
                    '<Default Extension="png" ContentType="image/png"/>'
                    '<Default Extension="rels"', 1)
    if "/xl/cellimages.xml" not in ct:
        ct = ct.replace("</Types>", '<Override PartName="/xl/cellimages.xml" '
                        'ContentType="application/vnd.wps-officedocument.cellimage+xml"/></Types>')
    if "cellImage" not in rels:
        rels = rels.replace("</Relationships>", '<Relationship Id="rId99" '
                            'Type="http://www.wps.cn/officeDocument/2020/cellImage" '
                            'Target="cellimages.xml"/></Relationships>')

    tmp = out + ".tmp"
    with zipfile.ZipFile(out) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = ct.encode()
            elif item.filename == "xl/_rels/workbook.xml.rels":
                data = rels.encode()
            zout.writestr(item, data)
        have = set(zin.namelist())
        for n, data in parts.items():
            if n not in have:
                zout.writestr(n, data)
    shutil.move(tmp, out)
    return out





if __name__ == "__main__":
    print("==================步骤一：读取商品配置表==================")
    products = read_products()
    # print(products.head())

    print("==================步骤二：读取物流运费表==================")
    rates = read_freight()
    # print(rates)

    print("==================步骤三：读取订单表==================")
    orders = read_orders()
    print(orders.head())

    print("==================步骤四：将国家拆分为和物流运费表中一致的内容==================")
    orders = fix_country(orders)
    # print(orders)

    print("==================步骤五：根据SKU前缀匹配商品配置表和订单数据表==================")
    match_orders = match_sku(orders, products)
    match_orders['商品数量'] = pd.to_numeric(match_orders['商品数量'], errors="coerce").fillna(1).astype(int)
    # print(match_orders['商品数量'].dtype)
    # print(match_orders['重量'].dtype)
    print(match_orders.head())
    # print(match_orders.columns)

    print("==================步骤六：根据国家和重量匹配对应的规则获取公斤价和和挂号费==================")
    match_orders = match_rate(match_orders, rates)
    # print(match_orders.head())
    # print(match_orders.columns)


    print("==================步骤七：计算附加费，毛利，欧盟国家费用等==================")
    detail = add_money(match_orders, service_rates(rates))
    # print(detail.head())

    issues = build_issues(detail, rates=rates)      # rates 用于 ROUTE_FALLBACK
    # path = write_excel(detail, issues, "输出/60103计算表.xlsx")
    # print("已生成:", path)