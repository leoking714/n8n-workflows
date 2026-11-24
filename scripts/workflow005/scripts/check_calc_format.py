#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ROLE: 计算表格式自检
ENTRY: python3 check_calc_format.py <xlsx>
CALLED-BY: write_from_template.py 写表后自动调用；也可手工跑
DEPS: 无

计算表格式自检：校验生成的 xlsx 是否与客户模板格式一致。
任何一项不通过 → 退出码 1（用于防止「合计块继承明细行样式」这类回归）。

用法：
  python check_calc_format.py <计算表.xlsx>
  python check_calc_format.py <计算表.xlsx> --verbose
"""
import argparse
import re
import sys
import zipfile

import openpyxl

TOTALS_STYLE = 1          # 模板里合计行/图例行的行样式索引（s=1，无底色）
LEGEND_HEIGHT = 18        # 图例行行高
LEGEND_FIRST_COLOR = "FFFFC7CE"   # 图例第 1 行 A 列（阻断·红）
LEGEND_CHIPS = {                  # 图例 5 个颜色块（A 列）
    "阻断": "FFFFC7CE", "警告": "FFFFEB9C", "欧盟": "FFDDEBF7",
    "偏远": "FFE4DFEC", "附加费": "FFFCE4D6",
}
EU_FILL = "FFDDEBF7"      # 欧盟订单行底色


def row_styles(path):
    """读 sheet1.xml 里的行级样式索引 {行号: s}"""
    out = {}
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if re.match(r"xl/worksheets/sheet\d+\.xml$", name):
                txt = z.read(name).decode("utf-8")
                for rn, s in re.findall(r'<row r="(\d+)"[^>]*?\ss="(\d+)"', txt):
                    out[int(rn)] = int(s)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--verbose", action="store_true")
    a = ap.parse_args()

    wb = openpyxl.load_workbook(a.xlsx)
    ws = wb.worksheets[0]
    styles = row_styles(a.xlsx)
    problems = []

    # 定位总计行
    tot = None
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").strip() == "总计":
            tot = r
            break
    if tot is None:
        problems.append("找不到“总计”行")
        print("\n".join("❌ " + p for p in problems))
        return 1

    # 1) 数据行：有内容、行高 54、样式 s=2
    data_rows = tot - 1
    if data_rows < 1:
        problems.append("总计行之前没有数据行")
    else:
        for r in range(2, tot):
            if styles.get(r) != 2:
                problems.append(f"数据行{r} 行样式={styles.get(r)}，应为 2")
                break
        if ws.row_dimensions[2].height not in (54, 54.0):
            problems.append(f"数据行行高={ws.row_dimensions[2].height}，应为 54")

    # 2) 合计块/图例：行样式必须 s=1（不能继承明细行的 s=2）
    legend_start = None
    for r in range(tot, ws.max_row + 1):
        if "颜色图例" in str(ws.cell(r, 2).value or ""):
            legend_start = r
            break
    end_row = ws.max_row
    for r in range(tot, end_row + 1):
        s = styles.get(r)
        if s != TOTALS_STYLE:
            problems.append(f"合计/图例行{r} 行样式={s}，应为 {TOTALS_STYLE}（否则底部会有淡色底纹）")

    # 3) 图例行高 = 18
    if legend_start:
        for r in range(legend_start + 1, end_row + 1):
            h = ws.row_dimensions[r].height
            if h not in (LEGEND_HEIGHT, float(LEGEND_HEIGHT)):
                problems.append(f"图例行{r} 行高={h}，应为 {LEGEND_HEIGHT}")
        # 4) 图例 5 个颜色块
        chips = {}
        for r in range(legend_start + 1, end_row + 1):
            f = ws.cell(r, 1).fill
            if f and f.fill_type:
                chips[r] = f.fgColor.rgb
        if not chips:
            problems.append("图例 A 列没有颜色块")
        for name, rgb in LEGEND_CHIPS.items():
            if rgb not in chips.values():
                problems.append(f"图例缺少颜色块：{name}（{rgb}）")
    else:
        problems.append("找不到“颜色图例”行")

    # 5) 合计公式
    if not str(ws.cell(tot, 32).value or "").startswith("=SUM"):
        problems.append(f"总计行 AF{tot} 不是 SUM 公式：{ws.cell(tot, 32).value!r}")

    # 6) 欧盟订单行应有蓝底
    eu_rows = [r for r in range(2, tot)
               if ws.cell(r, 22).value not in (None, "", 0)]
    for r in eu_rows:
        f = ws.cell(r, 1).fill
        if not (f and f.fill_type and f.fgColor.rgb == EU_FILL):
            problems.append(f"欧盟行{r} 缺少蓝色底纹（{EU_FILL}）")
            break

    if a.verbose:
        print(f"文件：{a.xlsx}")
        print(f"数据行 {data_rows} | 总计行 {tot} | 图例起始 {legend_start} | 总行数 {ws.max_row}")
        print(f"合计块行样式：{ {r: styles.get(r) for r in range(tot, end_row + 1)} }")

    if problems:
        print("\n".join("❌ " + p for p in problems))
        print(f"\n共 {len(problems)} 项不通过")
        return 1
    print(f"✅ 格式检查通过（数据行 {data_rows}，合计行 {tot}）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
