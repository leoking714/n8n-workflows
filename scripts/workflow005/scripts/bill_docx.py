#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
用客户 docx 模板生成 Bill：
  1) 读计算表 xlsx → 逐单数据（交易号/件数/国家/UHF/金额/图片）
  2) 填 docx：Bill Date / Customer Number / Prepared for / Account Manager / 明细表 / 合计
  3) docx → PDF（soffice）

用法：
  python bill_docx.py --xlsx <计算表.xlsx> [--out <bill.pdf>] [--date "2026. 11. 20"] [--cfg <json>]
cfg 里可带：customer_no / customer_name / account_manager
"""
import argparse, copy, datetime, json, os, re, subprocess, sys, tempfile

import openpyxl
from docx import Document
from docx.shared import Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bill_pdf_from_xlsx import load as load_xlsx, parse_weight, num   # 复用取数逻辑
import bill_pdf as B

TPL = "/data/data/workflow005/config/bill_template.docx"
CRED = "/data/data/workflow005/config/feishu_newbase.json"
PARAM_TABLE = "tblDPHP5ZgioWinV"


def load_cfg_from_feishu():
    """从「全局参数」表读 客户编号 / 客户名 / 业务员"""
    try:
        import urllib.request
        cfg = json.load(open(CRED, encoding="utf-8"))
        req = urllib.request.Request(
            "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
            data=json.dumps({"app_id": cfg["app_id"], "app_secret": cfg["app_secret"]}).encode(),
            headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=20) as r:
            token = json.loads(r.read().decode())["tenant_access_token"]
        url = (f"https://open.feishu.cn/open-apis/bitable/v1/apps/{cfg['base']}"
               f"/tables/{PARAM_TABLE}/records/search?page_size=200")
        req2 = urllib.request.Request(url, data=b'{"filter":{"conjunction":"and","conditions":[]}}',
                                      headers={"Content-Type": "application/json",
                                               "Authorization": "Bearer " + token})
        with urllib.request.urlopen(req2, timeout=20) as r:
            items = json.loads(r.read().decode())["data"]["items"]

        def flat(v):
            if isinstance(v, list):
                return "".join(str(x.get("text", "")) if isinstance(x, dict) else str(x) for x in v)
            if isinstance(v, dict):
                return str(v.get("text") or v.get("value") or "")
            return "" if v is None else str(v)

        P = {}
        for it in items:
            f = it.get("fields", {})
            k = flat(f.get("参数键")).strip()
            if k:
                P[k] = flat(f.get("参数值")).strip()
        return {"customer_no": P.get("CUSTOMER_NO", ""),
                "customer_name": P.get("CUSTOMER_NAME", ""),
                "account_manager": P.get("ACCOUNT_MANAGER", ""),
                "eu_topup": float(P.get("EU_TOPUP_PRICE") or 0)}
    except Exception as e:
        print("读全局参数失败（用默认）:", e, file=sys.stderr)
        return {}


def set_cell_keep_style(cell, text, size=None):
    """替换单元格文字，保留第一个 run 的格式（可指定字号）"""
    p = cell.paragraphs[0]
    if p.runs:
        run = p.runs[0]
        run.text = text
        for r in p.runs[1:]:
            r.text = ""
    else:
        run = p.add_run(text)
    if size:
        try:
            run.font.size = Pt(size)
        except Exception:
            pass
    try:
        run.font.name = "DejaVu Sans"
        run.font.bold = False
        from docx.oxml.ns import qn as _qn
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(_qn('w:rFonts'))
        if rFonts is not None:
            rFonts.set(_qn('w:ascii'), 'DejaVu Sans')
            rFonts.set(_qn('w:hAnsi'), 'DejaVu Sans')
            rFonts.set(_qn('w:eastAsia'), 'DejaVu Sans')
            rFonts.set(_qn('w:cs'), 'DejaVu Sans')
    except Exception:
        pass


def replace_in_paragraph(para, old, new):
    if old in para.text:
        # 整段重写，保留第一个 run 样式
        first = para.runs[0] if para.runs else para.add_run()
        text = para.text.replace(old, new)
        for r in para.runs[1:]:
            r._element.getparent().remove(r._element)
        first.text = text
        return True
    return False


def fill_header(doc, date_text, cfg):
    from docx.enum.table import WD_ROW_HEIGHT_RULE
    t0 = doc.tables[0]
    # 表头固定 144pt（和客户 Bill 的横幅等高），内部文字 8pt、段距 0
    try:
        for row in t0.rows:
            row.height = Pt(144)
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    except Exception:
        pass
    # 缩小 logo
    for cell in t0.rows[0].cells:
        for shape in cell._tc.xpath('.//pic:pic'):
            try:
                from docx.shared import Emu
                blip = shape.xpath('.//a:blip')[0]
                rid = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                # 通过 inline shape 调整
            except Exception:
                pass
    # 垂直居中 + 段距清零
    try:
        from docx.enum.table import WD_ALIGN_VERTICAL
        for row in t0.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    except Exception:
        pass
    for row in t0.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                try:
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after = Pt(0)
                except Exception:
                    pass
    for row in t0.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                txt = para.text
                if not txt.strip():
                    continue
                m = re.match(r"^Account Manager:\s*(.*)$", txt.split("\n")[0].strip())
                # 多行段落里逐行替换
                lines = txt.split("\n")
                changed = False
                for i, line in enumerate(lines):
                    s = line.strip()
                    if s.startswith("Account Manager:"):
                        lines[i] = re.sub(r"(Account Manager:\s*).*", r"\g<1>" + (cfg.get("account_manager") or ""), line)
                        changed = True
                    elif s.startswith("Bill Date:"):
                        lines[i] = re.sub(r"(Bill Date:\s*).*", r"\g<1>" + date_text, line)
                        changed = True
                    elif s.startswith("Customer Number:"):
                        lines[i] = re.sub(r"(Customer Number:\s*).*", r"\g<1>" + (cfg.get("customer_no") or ""), line)
                        changed = True
                    elif s.startswith("Prepared for:"):
                        lines[i] = re.sub(r"(Prepared for:\s*).*", r"\g<1>" + (cfg.get("customer_name") or ""), line)
                        changed = True
                if changed:
                    first = para.runs[0] if para.runs else para.add_run()
                    for r in para.runs[1:]:
                        r._element.getparent().remove(r._element)
                    first.text = "\n".join(lines)
                    try:
                        first.font.size = Pt(8)
                    except Exception:
                        pass


def set_row_image(cell, imgs, width_in=0.29):
    """把订单的所有商品图并排放进单元格（每张等比缩到 20.7x26pt 框内）"""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn as _qn
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.left_indent = None
    pf.first_line_indent = None
    pf.right_indent = None

    # 清掉模板里残留的文字和图片
    for para in cell.paragraphs:
        for r in list(para.runs):
            r._element.getparent().remove(r._element)
    for pic in cell._tc.xpath(".//pic:pic"):
        el = pic
        while el is not None and el.tag not in (
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing",
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pict"):
            el = el.getparent()
        if el is not None and el.getparent() is not None:
            el.getparent().remove(el)

    if isinstance(imgs, str):
        imgs = [imgs]
    imgs = [f for f in (imgs or []) if f and os.path.exists(f)]
    if not imgs:
        return

    from PIL import Image as PILImage
    box_w, box_h = 20.7 / 72.0, 26.0 / 72.0
    for idx, path in enumerate(imgs):
        try:
            if idx > 0:
                sp = p.add_run(" ")
                sp.font.size = Pt(3)
            run = p.add_run()
            with PILImage.open(path) as im:
                iw, ih = im.size
            scale = min(box_w / iw, box_h / ih)
            run.add_picture(path, width=Inches(iw * scale), height=Inches(ih * scale))
        except Exception:
            continue


def fill_table(doc, rows, date_text, eu_topup=0.0, last_balance_row=0.0):
    t = doc.tables[1]
    # 列宽对齐客户 Bill（pt）：No./Order/item/qty/country/UHF/Amount
    from docx.shared import Pt as _Pt
    widths = [30.0, 66.0, 216.0, 46.0, 76.0, 54.0, 53.0]
    try:
        from docx.oxml.ns import qn as _qn
        from docx.oxml import OxmlElement as _El
        tbl = t._tbl
        tblPr = tbl.tblPr
        # 固定布局
        layout = tblPr.find(_qn('w:tblLayout'))
        if layout is None:
            layout = _El('w:tblLayout'); tblPr.append(layout)
        layout.set(_qn('w:type'), 'fixed')
        # 表格总宽
        tblW = tblPr.find(_qn('w:tblW'))
        if tblW is None:
            tblW = _El('w:tblW'); tblPr.append(tblW)
        tblW.set(_qn('w:type'), 'dxa'); tblW.set(_qn('w:w'), str(int(sum(widths) * 20)))
        # gridCol
        grid = tbl.find(_qn('w:tblGrid'))
        if grid is not None:
            for gc in list(grid):
                grid.remove(gc)
            for w in widths:
                gc = _El('w:gridCol'); gc.set(_qn('w:w'), str(int(w * 20))); grid.append(gc)
        # 每格 tcW
        for row in t.rows:
            for i, cell in enumerate(row.cells):
                tcPr = cell._tc.get_or_add_tcPr()
                tcW = tcPr.find(_qn('w:tcW'))
                if tcW is None:
                    tcW = _El('w:tcW'); tcPr.append(tcW)
                tcW.set(_qn('w:type'), 'dxa'); tcW.set(_qn('w:w'), str(int(widths[i] * 20)))
    except Exception as e:
        print('列宽设置失败', e)
    # 单元格内边距 3pt（客户 Bill 的边距很小，否则表头会换行）
    try:
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        tblPr = t._tbl.tblPr
        mar = tblPr.find(qn('w:tblCellMar'))
        if mar is None:
            mar = OxmlElement('w:tblCellMar')
            tblPr.append(mar)
        for side in ('left', 'right', 'top', 'bottom'):
            el = mar.find(qn(f'w:{side}'))
            if el is None:
                el = OxmlElement(f'w:{side}')
                mar.append(el)
            el.set(qn('w:w'), '40')
            el.set(qn('w:type'), 'dxa')
    except Exception:
        pass

    # 去掉整表的加粗（客户 Bill 正文不加粗）
    try:
        from docx.oxml.ns import qn as _qn2
        for row in t.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.bold = False
                    pPr = para._p.get_or_add_pPr()
                    for b in pPr.findall(_qn2('w:b')):
                        pPr.remove(b)
    except Exception:
        pass

    # 表头行：字号/换行与客户版一致
    try:
        hdr = t.rows[0]
        hdr.height = _Pt(22)
        cells = hdr.cells
        heads = ["No.", "Order number", "item name", "quantity", "country", "Union Handling\nFee (UHF)", "Amount"]
        for i, cell in enumerate(cells):
            para = cell.paragraphs[0]
            if para.runs:
                run = para.runs[0]
                for extra in para.runs[1:]:
                    extra._element.getparent().remove(extra._element)
            else:
                run = para.add_run()
            run.text = heads[i]
            run.font.size = _Pt(6.5 if i == 5 else 8)
            run.font.name = "DejaVu Sans"
        run.font.bold = False
    except Exception:
        pass
    need = len(rows)
    have = len(t.rows) - 1

    # 行数调整
    if need > have:
        last = t.rows[-1]._tr
        for _ in range(need - have):
            t._tbl.append(copy.deepcopy(last))
    elif need < have:
        for _ in range(have - need):
            tr = t.rows[-1]._tr
            tr.getparent().remove(tr)

    # 末尾补一行 Last Balance（只做展示，不计入 Total）
    total_rows = need + 1
    have2 = len(t.rows) - 1
    if total_rows > have2:
        last = t.rows[-1]._tr
        for _ in range(total_rows - have2):
            t._tbl.append(copy.deepcopy(last))
    elif total_rows < have2:
        for _ in range(have2 - total_rows):
            tr = t.rows[-1]._tr
            tr.getparent().remove(tr)

    for i, r in enumerate(rows):
        row = t.rows[i + 1]
        cells = row.cells
        set_cell_keep_style(cells[0], str(i + 1), 8)
        set_cell_keep_style(cells[1], str(r["txn"]), 8)
        set_row_image(cells[2], r.get("imgs") or [])
        set_cell_keep_style(cells[3], str(r["qty"]), 8)
        set_cell_keep_style(cells[4], str(r["country"]), 8)
        # UHF 列：欧盟单显示补收金额（默认 2.3）
        uhf = eu_topup if (r.get("uhf") or 0) > 0 and eu_topup > 0 else (r.get("uhf") or 0)
        set_cell_keep_style(cells[5], "US$%.2f" % uhf if uhf else "/", 7)
        set_cell_keep_style(cells[6], "US$%.2f" % r["amount"], 8)

    # 末尾 Last Balance 行（展示用，不计入 Total）
    try:
        lb_row = t.rows[len(rows) + 1]
        lc = lb_row.cells
        set_cell_keep_style(lc[0], str(len(rows) + 1), 8)
        set_cell_keep_style(lc[1], "Last Balance", 8)
        set_row_image(lc[2], [])
        try:
            from docx.enum.text import WD_ALIGN_PARAGRAPH as _AL
            pp = lc[2].paragraphs[0]
            pp.alignment = _AL.CENTER
            r = pp.add_run("Last Balance")
            r.font.size = Pt(8)
            r.font.name = "DejaVu Sans"
            r.font.bold = False
        except Exception:
            pass
        set_cell_keep_style(lc[3], "", 8)
        set_cell_keep_style(lc[4], "--", 8)
        set_cell_keep_style(lc[5], "/", 7)
        set_cell_keep_style(lc[6], "US$%.2f" % (last_balance_row or 0), 8)
    except Exception as e:
        print("Last Balance 行填充失败", e)


def fill_totals(doc, subtotal, last_balance=0.0):
    fee = round(subtotal * 0.005, 2)
    total = round(subtotal + fee + (last_balance or 0), 2)

    t = doc.tables[2]
    # 去掉空白行（模板用空行做行距）
    for row in list(t.rows):
        if not row.cells[0].text.strip():
            row._tr.getparent().remove(row._tr)

    labels = ["Total($):", "payment fee(0.5%):", "Last Balance:", "Total($):"]
    values = [subtotal, fee, last_balance or 0, total]
    rows = list(t.rows)
    # 需要 4 行，不够就补
    while len(rows) < 4:
        tr = rows[-1]._tr
        t._tbl.append(copy.deepcopy(tr))
        rows = list(t.rows)
    for i in range(4):
        p = rows[i].cells[0].paragraphs[0]
        text = f"{labels[i]}\tUS${values[i]:,.2f}"
        if p.runs:
            p.runs[0].text = text
            for r in p.runs[1:]:
                r._element.getparent().remove(r._element)
        else:
            p.add_run(text)
    return subtotal, fee, total


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--out", default="")
    ap.add_argument("--date", default="")
    ap.add_argument("--cfg", default="{}")
    ap.add_argument("--last-balance", default="0")
    ap.add_argument("--tpl", default=TPL)
    a = ap.parse_args()

    cfg = load_cfg_from_feishu()
    cfg.update({k: v for k, v in json.loads(a.cfg or "{}").items() if v})
    rows = load_xlsx(a.xlsx)
    if not rows:
        raise SystemExit("计算表里没有数据行")

    date_text = a.date or datetime.datetime.now().strftime("%Y. %-m. %-d")
    out_pdf = a.out or (os.path.splitext(a.xlsx)[0] + "_docx.pdf")

    doc = Document(a.tpl)
    # 页边距与客户 Bill 对齐（顶部 0，让横幅贴顶）
    try:
        for sec in doc.sections:
            sec.top_margin = Pt(36)
            sec.bottom_margin = Pt(36)
            sec.left_margin = Pt(27)
            sec.right_margin = Pt(27)
    except Exception:
        pass
    fill_header(doc, date_text, cfg)
    eu_topup = float(cfg.get('eu_topup') or 0)
    try:
        lb = float(a.last_balance or 0)
    except Exception:
        lb = 0.0
    fill_table(doc, rows, date_text, eu_topup, lb)
    # 合计 = 各单金额 + 欧盟单 × 2.3（UHF 列金额并入合计）
    eu_cnt = len([r for r in rows if (r.get("uhf") or 0) > 0])
    subtotal = round(sum(r["amount"] or 0 for r in rows) + eu_cnt * eu_topup, 2)
    st, fee, total = fill_totals(doc, subtotal, lb)

    tmpdir = tempfile.mkdtemp(prefix="billdocx_")
    docx_path = os.path.join(tmpdir, "bill.docx")
    doc.save(docx_path)

    r = subprocess.run(["soffice", "--headless", "--convert-to", "pdf",
                        "--outdir", tmpdir, docx_path],
                       capture_output=True, text=True, timeout=600)
    pdf_tmp = os.path.join(tmpdir, "bill.pdf")
    if not os.path.exists(pdf_tmp):
        print(r.stderr[-800:], file=sys.stderr)
        raise SystemExit("docx → pdf 失败")
    # 跨设备（/tmp → /data）不能用 os.replace，统一 copy
    os.makedirs(os.path.dirname(out_pdf), exist_ok=True)
    subprocess.run(["cp", "-f", pdf_tmp, out_pdf], check=True)

    print(json.dumps({"ok": True, "out": out_pdf, "orders": len(rows),
                      "subtotal": st, "fee": fee, "last_balance": lb, "total": total},
                     ensure_ascii=False))


if __name__ == "__main__":
    main()
