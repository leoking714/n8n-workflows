# -*- coding: utf-8 -*-
"""
60103 计算表 · pandas 独立版

【这个文件不依赖 生成计算表.py，单独拿走就能跑】
pandas 负责「读表 + 清洗 + 配对」，openpyxl 负责「按模板写 Excel」。

依赖：pandas、openpyxl（不用 pyyaml）
运行：python 生成计算表_pandas.py
输出：输出/60103计算表_pandas_<时间戳>.xlsx

目录约定（与 openpyxl 版相同）：
    源文件/订单导出数据.xlsx     ← 每行一件商品，订单字段只在首行
    源文件/物流运费.xlsx         ← 国家 + 线路 + 重量段 → 公斤价/挂号费
    源文件/商品配置.xlsx         ← SKU 前缀 → 采购价/国内运费/重量
    源文件/计算表模板.xlsx       ← 必须以它为基础改，样式/列宽/商品图都在里面
    输出/                        ← 目录不存在会自动创建

与 openpyxl 版的区别：只在于「读表 + 计算」用 DataFrame 做，
写入 Excel 的那一半是同一份逻辑（为避免两处维护，本文件内联了一份）。
"""

import datetime
import os
import re
import shutil
import zipfile
from copy import copy

import openpyxl
import pandas as pd
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter


pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)


# ============================ 0. 路径与口径 ============================
# 以「脚本所在目录」为根，这样不管在哪个 cwd 下运行（n8n / 定时任务）都能找到文件
BASE = os.environ.get("BILL_BASE") or os.path.dirname(os.path.abspath(__file__))
S = os.path.join(BASE, "data")

ORDER = os.path.join(S, "订单导出数据20261104182737-623523-652214512317620224.xlsx")
FREIGHT = os.path.join(S, "60103物流运费-11.2更.xlsx")
PRODUCT = os.path.join(S, "商品配置-11.2更.xlsx")
TEMPLATE = os.path.join(S, "60103计算表_20261104_183110.xlsx")

NOW = datetime.datetime.now()
OUT_DIR = os.environ.get("BILL_OUT") or os.path.join(BASE, "输出")
os.makedirs(OUT_DIR, exist_ok=True)          # ★ 目录不存在就建：shutil.copy 不会自动建目录
OUT = os.path.join(OUT_DIR, f"60103计算表_pandas_{NOW:%Y%m%d_%H%M%S}.xlsx")

for _f in (ORDER, FREIGHT, PRODUCT, TEMPLATE):          # 缺文件时报一句人话，而不是 traceback
    if not os.path.exists(_f):
        raise SystemExit(f"找不到输入文件：{_f}")

FX = 6.65                                  # 美元汇率：AC = AA / FX
MULT = 1.15                                # 服务费：AA 先 ×1.15
EUR, EUR2USD, EU_TAX = 3, 0.9, 3.33        # 欧盟税三列 T / U / V
EU = {"德国", "法国", "意大利", "西班牙", "荷兰", "比利时", "卢森堡", "奥地利",
      "葡萄牙", "爱尔兰", "丹麦", "瑞典", "芬兰", "希腊", "波兰", "捷克", "匈牙利",
      "罗马尼亚", "保加利亚", "克罗地亚", "斯洛文尼亚", "斯洛伐克", "爱沙尼亚",
      "拉脱维亚", "立陶宛", "马耳他", "塞浦路斯"}       # 英国已脱欧
SUR = {"墨西哥": 2.9, "智利": 7, "卡塔尔": 41, "沙特": 40, "阿联酋": 2}
FILL = {"blocked": "FFFFC7CE", "eu": "FFDDEBF7", "remote": "FFE4DFEC", "surcharge": "FFFCE4D6"}

# 欧盟税三列自洽提示：金额只认 EU_TAX（V 列），T/U 只是留档
if abs(EUR / EUR2USD - EU_TAX) > 0.01:
    print(f"⚠️ 欧盟税三列不一致：T/U = {EUR}/{EUR2USD} → {EUR / EUR2USD:.2f}，但 V = {EU_TAX}"
          f"（金额按 V 算，T/U 会显示成另一套）")


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

    print(df.head())
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


# ============================ 3. 写进计算表 ============================
LEGEND = ["颜色图例：",
          "阻断：缺运价/缺配置等，金额未计（明细格见「未查询」）",
          "警告：需人工确认（原因见异常表）",
          f"欧盟国家：收欧盟税 {EU_TAX}$/单（明细 T/U/V 列）",
          "美国（偏远）：专线改写并套偏远列价格",
          "含附加费国家（墨西哥/智利/卡塔尔/沙特/阿联酋）",
          "金额格「未查询」= 运价表中未查到该运价"]


def fill_rows(ws, rows):
    """
    把二维列表覆盖写入第 2 行起（用于「汇总」「异常」这种纯值表）。

    输入：工作表、rows（每行是一个列表）
    输出：无（原地改工作表）
    注意：模板行数比数据多时删掉多余旧行；数据比模板多时新行套第 2 行样式，
          否则新行没有边框（这是“看起来不齐”的常见原因）
    """
    old = ws.max_row
    for i, row in enumerate(rows):
        for j, v in enumerate(row, 1):
            c = ws.cell(row=2 + i, column=j, value=v)
            if 2 + i > old:
                c._style = copy(ws.cell(row=2, column=j)._style)
    if old > 1 + len(rows):
        ws.delete_rows(2 + len(rows), old - 1 - len(rows))


def write(result, issues, out=None, now=None):
    """
    按模板写出一份新的计算表，返回文件路径。

    输入：calc() 的 result / issues；out 输出路径（默认用文件顶部的 OUT）
    流程：复制模板 → 记住第 2 行样式和 SKU→图片映射 → 清空旧数据行 →
          逐行写明细（值/公式/合并/底色）→ 总计行 → 图例 → 汇总 → 异常 →
          配置快照 → 存盘 → 补回 WPS 图片部件
    注意：必须“复制模板再改”，不能新建 Workbook —— 否则样式、列宽、合并、
          商品图全部丢失，这也是为什么不直接用 pandas 的 to_excel
    """
    out = out or OUT
    now = now or NOW
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    shutil.copy(TEMPLATE, out)                       # 以模板为底：样式/列宽/图片都在里面
    wb = openpyxl.load_workbook(out)
    ws = wb.worksheets[0]
    ws.title = f"{now:%m.%d}计算表"

    cols = [get_column_letter(i) for i in range(1, 39)]
    style = {c: copy(ws[f"{c}2"]._style) for c in cols}          # 记住数据行样式
    height = ws.row_dimensions[2].height
    img = {row[9].value: row[6].value                            # SKU → 图片公式
           for row in ws.iter_rows(min_row=2) if row[9].value and row[6].value}

    for rng in list(ws.merged_cells.ranges):                     # 先拆旧合并
        ws.unmerge_cells(str(rng))
    ws.delete_rows(2, ws.max_row)                                # 清空旧内容，只留表头

    r = 2
    for g in result:
        first, last = r, r + len(g["rows"]) - 1
        for i, (qty, sku, buy, dom, wt, rmb, usd) in enumerate(g["rows"]):
            v = {"A": "", "G": img.get(sku, ""), "H": qty, "J": sku, "K": buy, "L": dom,
                 "M": "万邦普货", "O": "带", "Z": MULT, "AB": FX, "AC": f"=AA{r}/AB{r}"}
            if i == 0:                                            # 订单级字段只写首行
                v.update({"B": g["o"]["logi"], "C": g["o"]["chan"], "D": g["o"]["inner"],
                          "E": g["o"]["no"], "F": g["o"]["txn"], "I": g["country"],
                          "P": g["price"] if not g["blocked"] else "未查询",
                          "Q": g["fee"] if not g["blocked"] else "未查询",
                          "R": "=" + "+".join(f"H{first + k}*{row[4]}"
                                              for k, row in enumerate(g["rows"])),
                          "S": f"=N(P{r})*N(R{r})+N(Q{r})",
                          "AA": f"=(N(S{r})+N(K{r})*N(H{r})+N(L{r})*N(H{r})+N(Y{r}))*N(Z{r})",
                          "AE": f"=SUM(AC{first}:AC{last})",
                          "AD": f"=N(V{r})+AE{r}", "AF": f"=ROUND(AD{r},2)"})
                if g["tax"]:
                    v.update({"T": EUR, "U": EUR2USD, "V": g["tax"], "W": f"=V{r}*AB{r}"})
                if g["sur"]:
                    v["Y"] = g["sur"]
                note = list(g["notes"])
                if g["blocked"]:
                    note.append(f"阻断：「{g['country']}」缺 {g['weight']}kg 运价")
                if note:
                    v["AL"] = "；".join(note)
            else:
                v["AA"] = f"=(N(K{r})*N(H{r})+N(L{r})*N(H{r}))*N(Z{r})"

            for c in cols:
                cell = ws[f"{c}{r}"]
                cell._style = copy(style[c])
                if c in v:
                    cell.value = v[c]
                cell.fill = (PatternFill("solid", fgColor=FILL[cat_of(g)]) if cat_of(g)
                             else PatternFill(fill_type=None))
            ws.row_dimensions[r].height = height
            r += 1
        if last > first:                                          # 国家/重量/账单价格按单合并
            for c in ("I", "R", "AF"):
                ws.merge_cells(f"{c}{first}:{c}{last}")

    ws[f"A{r}"], ws[f"V{r}"] = "总计", f"=SUM(V2:V{r - 1})"
    ws[f"AE{r}"], ws[f"AF{r}"] = f"=SUM(AE2:AE{r - 1})", f"=SUM(AF2:AF{r - 1})"
    for i, text in enumerate(LEGEND):
        ws[f"B{r + 3 + i}"] = text

    # ---- 汇总：按国家小计（阻断单不计）----
    agg = {}
    for g in result:
        if g["blocked"]:
            continue
        a = agg.setdefault(g["country"], dict(o=set(), qty=0, wt=0, intl=0,
                                              dom=0, buy=0, sur=0, eu=0, usd=0))
        a["o"].add(g["o"]["txn"])
        a["qty"] += sum(x[0] for x in g["rows"])
        a["wt"] += g["weight"]
        a["intl"] += g["price"] * g["weight"] + g["fee"]
        a["sur"] += g["sur"]
        a["eu"] += round(g["tax"] * FX, 2)
        a["usd"] += g["bill"]
        for qty, _sku, buy, dom, _wt, _rmb, _usd in g["rows"]:
            a["buy"] += buy * qty
            a["dom"] += dom * qty
    rows = [[k, len(a["o"]), a["qty"], round(a["wt"], 2), round(a["intl"], 2),
             round(a["dom"], 2), round(a["buy"], 2), round(a["sur"], 2),
             round(a["eu"], 2), round(a["usd"], 2)] for k, a in agg.items()]
    rows.append(["合计"] + [round(sum(x[j] for x in rows), 2) for j in range(1, 10)])
    fill_rows(wb["汇总"], rows)

    # ---- 异常：阻断在前，警告在后 ----
    fill_rows(wb["异常"], [[lv, txn, no, src, "", msg]
                           for lv, txn, no, src, msg in
                           sorted(issues, key=lambda x: x[0] != "阻断")])

    # ---- 配置快照：更新本次跑批信息 ----
    snap = wb["配置快照"]
    for key, val in {"批次时间": f"{NOW:%Y-%m-%d %H:%M:%S}", "订单源文件": ORDER,
                     "美元汇率": str(FX), "欧盟税费($)": f"{EU_TAX}（25号后5.5条款已忽略）"}.items():
        for i in range(2, snap.max_row + 1):
            if snap[f"A{i}"].value == key:
                snap[f"B{i}"] = val

    wb.save(out)
    keep_wps_images(TEMPLATE, out)                   # 补回 WPS 单元格图片
    return out


def cat_of(g):
    """
    判断一个订单该刷什么底色。

    输入：calc() 里的一个订单 dict
    输出：'blocked' / 'eu' / 'remote' / 'surcharge' / ''（无底色）之一
    规则：优先级 阻断 > 欧盟 > 偏远 > 附加费，一行只刷一种
    """
    if g["blocked"]:
        return "blocked"
    if g["tax"]:
        return "eu"
    if g["country"] == "美国（偏远）":
        return "remote"
    if g["sur"]:
        return "surcharge"
    return ""


def keep_wps_images(tpl, out):
    """
    把 WPS 的单元格图片部件从模板补回新文件（原地改 zip）。

    背景：G 列是 =DISPIMG("ID_xxx",1)，图片不在 drawing 里，而在
          xl/cellimages.xml + xl/media/*.png，用 WPS 私有扩展挂在 workbook.xml.rels 上；
          openpyxl 存盘会把这几样一起丢掉，图片就全空了。
    输入：模板路径、刚写好的文件路径
    输出：无（直接覆盖 out）
    注意：① 图片部件从模板取，但 [Content_Types].xml 和 workbook.xml.rels 必须
            读“新文件”的 —— 模板那份声明了 sharedStrings，新文件没有，弄错会打不开；
          ② 还要补 png 的 Default 声明和 cellImage 关系，否则 WPS 不认图
    """
    with zipfile.ZipFile(tpl) as z:                       # 图片部件从模板取
        parts = {n: z.read(n) for n in z.namelist()
                 if n in ("xl/cellimages.xml", "xl/_rels/cellimages.xml.rels")
                 or n.startswith("xl/media/")}
    with zipfile.ZipFile(out) as z:                       # 目录文件必须用新文件的
        ct = z.read("[Content_Types].xml").decode()
        rels = z.read("xl/_rels/workbook.xml.rels").decode()
    ct = ct.replace("<Default Extension=\"rels\"",
                    '<Default Extension="png" ContentType="image/png"/>'
                    '<Default Extension="rels"', 1)
    if "/xl/cellimages.xml" not in ct:
        ct = ct.replace("</Types>", '<Override PartName="/xl/cellimages.xml" ContentType='
                        '"application/vnd.wps-officedocument.cellimage+xml"/></Types>')
    if "cellImage" not in rels:
        rels = rels.replace("</Relationships>", '<Relationship Id="rId99" Type="http://www.wps.cn'
                            '/officeDocument/2020/cellImage" Target="cellimages.xml"/></Relationships>')

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


if __name__ == "__main__":
    products = read_products()
    # print("=====================第一步=====================")
    # print(products.head())
    rates = read_freight()
    # print("=====================第二步=====================")
    # print(rates.head())

    orders = fix_country(read_orders())                 # ① 读订单 + 改写国家
    # print("=====================第三步=====================")
    # print(orders.head())

    match_df = match_sku(orders, products)
    # print("=====================第四步=====================")
    # print(match_df.head())


    detail = calc(match_df, rates)   # ② 配价格 → ③ 算金额
    # print("=====================第五步=====================")
    # print(detail)
    result, issues = to_result(detail)                  # ④ DataFrame → 每单一 dict
    print(result)
    print(issues)
    path = write(result, issues)                        # ⑤ 按模板写 Excel

    print(f"✔ {path}")
    print(f"  {len(result)} 单 / {len(detail)} 行 / 阻断 {sum(g['blocked'] for g in result)} 单")
    for level, txn, no, _src, msg in issues:
        print(f"  [{level}] {txn} {no} {msg}")
