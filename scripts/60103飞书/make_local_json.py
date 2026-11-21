# -*- coding: utf-8 -*-
"""
开发用：从本地 Excel 生成 data/*.json，模拟「飞书多维表格」拉下来的数据。

用途：没有飞书凭据时也能离线跑通计算流程、做回归对比。
线上不用这个脚本，线上用 sync_config.py 从飞书拉。

用法：
    python make_local_json.py
"""
import json
import os
import re

import openpyxl
import pandas as pd

import paths

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = paths.data_dir()
SRC = os.path.join(BASE, "示例数据")          # 放 5 个导入用 Excel


def _num(x):
    try:
        v = float(x)
        return v if v == v else None
    except (TypeError, ValueError):
        return None


def products():
    f = os.path.join(SRC, "1-商品配置.xlsx")
    df = pd.read_excel(f, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    rows = []
    for _, r in df.iterrows():
        sku = str(r.get("SKU前缀") or "").strip()
        if not sku:
            continue
        rows.append({
            "SKU前缀": sku,
            "采购价": _num(r.get("采购价")),
            "偏远采购价": _num(r.get("偏远采购价")),
            "国内运费": _num(r.get("国内运费")),
            "偏远国内运费": _num(r.get("偏远国内运费")),
            "重量": _num(r.get("重量")),
            "服务费_美国": _num(r.get("服务费_美国")),
            "服务费": _num(r.get("服务费")),
            "运价覆盖_国家": (str(r["运价覆盖_国家"]).strip()
                             if pd.notna(r.get("运价覆盖_国家")) else None),
            "运价覆盖_公斤价": _num(r.get("运价覆盖_公斤价")),
            "运价覆盖_挂号费": _num(r.get("运价覆盖_挂号费")),
            "账单固定金额": _num(r.get("账单固定金额")),
            "每计费单位件数": _num(r.get("每计费单位件数")) or 1,
            "状态": (str(r.get("状态") or "已确认").strip()),
        })
    return rows


def rates():
    f = os.path.join(SRC, "2-运价表.xlsx")
    df = pd.read_excel(f, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    rows = []
    for _, r in df.iterrows():
        c = str(r.get("国家") or "").strip()
        if not c:
            continue
        rows.append({
            "国家": c,
            "线路": str(r.get("线路") or "").strip(),
            "下界": _num(r.get("下界")),
            "上界": _num(r.get("上界")),
            "公斤价": _num(r.get("公斤价")),
            "挂号费": _num(r.get("挂号费")),
            "毛利": _num(r.get("毛利")),
            "附加费": _num(r.get("附加费")),
            "税费": _num(r.get("税费")),
        })
    return rows


def au_zones():
    f = os.path.join(SRC, "3-澳洲分区.xlsx")
    df = pd.read_excel(f, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    rows = []
    for _, r in df.iterrows():
        p = str(r.get("邮编") or "").strip().zfill(4)
        z = _num(r.get("分区"))
        if len(p) == 4 and z is not None:
            rows.append({"邮编": p, "分区": int(z)})
    return rows


def params():
    f = os.path.join(SRC, "4-全局参数.xlsx")
    df = pd.read_excel(f, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    rows = []
    for _, r in df.iterrows():
        k = str(r.get("参数键") or "").strip()
        if k:
            rows.append({"参数键": k, "参数值": str(r.get("参数值")),
                         "说明": str(r.get("说明") or "")})
    return rows


def country_rules():
    f = os.path.join(SRC, "5-国家规则.xlsx")
    df = pd.read_excel(f, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    rows = []
    for _, r in df.iterrows():
        c = str(r.get("国家") or "").strip()
        if not c:
            continue
        rows.append({
            "国家": c,
            "有运价": str(r.get("有运价") or "").strip(),
            "是否欧盟": str(r.get("是否欧盟") or "否").strip(),
            "订单附加费": _num(r.get("订单附加费")) or 0.0,
            "附加费类型": (str(r.get("附加费类型")).strip()
                           if pd.notna(r.get("附加费类型")) else None),
            "欧盟税额$": _num(r.get("欧盟税额$")),
            "服务费覆盖": _num(r.get("服务费覆盖")),
            "生效日期": (str(r.get("生效日期"))[:10]
                         if pd.notna(r.get("生效日期")) else None),
            "状态": str(r.get("状态") or "生效").strip(),
            "备注": (str(r.get("备注")).strip() if pd.notna(r.get("备注")) else None),
        })
    return rows


def main():
    os.makedirs(DATA, exist_ok=True)
    data = {
        "商品配置": products(),
        "运价表": rates(),
        "澳洲分区": au_zones(),
        "全局参数": params(),
        "国家规则": country_rules(),
    }
    for name, rows in data.items():
        with open(os.path.join(DATA, f"{name}.json"), "w", encoding="utf-8") as f:
            json.dump(rows, f, ensure_ascii=False, indent=2)
        print(f"  {name:8s} {len(rows):>5} 行")


if __name__ == "__main__":
    main()
