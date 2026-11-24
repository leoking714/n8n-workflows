# 60103 账单脚本（唯一生效版本）

> **只改这个目录里的文件。** 其它位置的 `bill_docx.py` / `write_from_template.py` 等副本已归档到
> `automation/scripts/_archive_20261124/`，不要再引用或修改它们。
>
> 容器内挂载路径：`/scripts/workflow005/scripts/`（n8n 节点里写的就是这个路径）

## 文件与调用关系

```
n8n 节点
 ├─ write_from_template.py     ← 计算表生成（workflow005 抓包版 / API 版）
 │    ├─ paths.py                  路径解析（/scripts 只读、/data 可写）
 │    └─ check_calc_format.py      写表后自检（子进程调用，失败只告警）
 │
 ├─ write_orders_xlsx.py       ← 订单 Excel（同两个 workflow005）
 │
 ├─ pick_and_render.py         ← Bill PDF 入口（workflow007）
 │    └─ bill_docx.py              套 docx 模板出 PDF（当前使用）
 │         ├─ bill_pdf_from_xlsx.py    xlsx → 逐单数据（件数/金额/图片）
 │         │    └─ bill_pdf.py             国家码/文案表 + 旧 reportlab 引擎（只当数据模块用）
 │         └─ check_calc_format.py
 │
 └─ sync_country_table.py      ← 国家表同步（workflow008）
```

## 各脚本职责

| 文件 | 作用 | 备注 |
|---|---|---|
| `write_from_template.py` | 把明细 JSON 套模板写成计算表 xlsx | 只排版，不算钱；金额由 n8n ④ 节点算好 |
| `check_calc_format.py` | 计算表格式自检（行样式/行高/图例/色块） | 退出码 1 表示格式回退 |
| `bill_docx.py` | **当前** Bill PDF 生成：填 docx → soffice → 后处理横幅/日期 | 顶部对齐、OpenAction、UHF 都在这里 |
| `bill_pdf_from_xlsx.py` | 从计算表读出逐单数据（含图片、金额重算） | `load()` 是入口 |
| `bill_pdf.py` | 旧版 reportlab 引擎 + 国家码/文案常量 | **别直接用它出 PDF**，只 `import` 常量 |
| `pick_and_render.py` | 选计算表文件 + 调 bill_docx | workflow007 用，`--engine docx` |
| `sync_country_table.py` | 运价表国家 → 国家表补齐（服务费 1.15/1.11） | 手动或 workflow008 调用 |
| `billing_ledger.py` | 账务层：preview（历史挂账）/ commit（写账单期·挂账·报价历史） | 出表前后调用 |
| `margin_report.py` | 毛利分析报表（汇总/单品多品/同款历史/逐单） | 出表后调用 |
| `feishu_send_file.py` | 上传文件到飞书群（xlsx/pdf） | workflow005「发送计算表文件」、workflow007「发送PDF文件」 |
| `write_orders_xlsx.py` | 订单明细导出 Excel | 供人工核对；依赖 `export_orders.py` |
| `export_orders.py` | 订单 JSON/XLSX 写出工具 | 被 `write_orders_xlsx.py` import，**不能删** |
| `paths.py` | 路径解析（容器/本地） | 被 write_from_template 引用 |

## 改代码前先确认

1. **改哪个文件？** 只有上面表格里的文件有效，`_archive_20261124/` 里的都是历史版本
2. **影响谁？** 改 `bill_docx.py` 影响所有 PDF；改 `write_from_template.py` 影响所有计算表
3. **改完验证：**
   ```bash
   # 计算表格式
   docker exec n8n-python-runner-1 python3 /scripts/workflow005/scripts/check_calc_format.py <xlsx>
   # 出表 + 自检（自动）
   docker exec n8n-python-runner-1 python3 /scripts/workflow005/scripts/write_from_template.py \
       --detail <detail.json> --out <out.xlsx>
   # 出 PDF
   docker exec n8n-python-runner-1 python3 /scripts/workflow005/scripts/bill_docx.py \
       --xlsx <xlsx> --out <pdf> --date '2026. 11. 24'
   ```

## 配置来源

- 飞书 Base：`config/feishu_newbase.json`（app_id/base/table id）
- 计算表模板：`config/计算表模板.xlsx`
- Bill 模板：`config/bill_template.docx`
- Bill 顶部横幅：`config/bill_header.png`（3571×864 = 6px/pt）
- 马帮 cookie：`config/mabang_cookie.txt`（仅独立 skill 用，n8n 流程走 data table）

业务规则（拆分单 `_1`、UHF、上期未付、quantity 口径等）见项目根目录 `AGENTS.md`。
