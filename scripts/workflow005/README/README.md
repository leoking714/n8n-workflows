# workflow005 · 订单匹配出账单

## 目录约定

```
代码（容器内只读挂载 /scripts）
/Users/zuo/n8n/files/automation/scripts/workflow005/
└── scripts/
    ├── export_orders.py          # 从 ERP 导出订单样本（JSON/XLSX）
    ├── write_from_template.py    # 明细 JSON → 计算表 xlsx + PDF
    └── paths.py                  # 路径解析

数据（容器内可写挂载 /data）
/Users/zuo/n8n/files/data/workflow005/
├── config/                        # 模板等配置
│   └── 计算表模板.xlsx
├── input/                         # 只放订单 Excel（给人看）
│   ├── 订单导出_YYYYMMDDHHMMSS.xlsx   ← 每次执行自动生成
│   └── orders_YYYYMMDD_YYYYMMDD.xlsx  ← 手工导出的样本
├── output/                        # 只放账单
│   ├── 60103计算表_YYYYMMDDHHMMSS.xlsx
│   └── 60103计算表_YYYYMMDDHHMMSS.pdf
└── logs/                          # 中间文件（JSON，给程序吃，可随时删）
    ├── orders_run_YYYYMMDDHHMMSS.json
    └── bill_detail_YYYYMMDDHHMMSS.json
```

## n8n 工作流

名称：**workflow005 · 订单匹配出账单**（31 节点）

```
触发 → 5 张飞书配置表（商品配置/运费表/澳洲分区/国家配置/全局参数）
     → 定义时间 → 循环拉 ERP 订单 → 解析出需要的数据
     → 提取澳洲邮编 → 查询澳大利亚分区表（按邮编单条查）→ 合并澳洲分区
     → 匹配商品配置（含完整性/运价校验）
     → 有缺失？
          ├ 有 → 飞书群消息（分 6 段列问题）
          └ 无 → 整理待写数据 → 保存明细JSON → 生成计算表(runner) → 解析出表结果
```

## 导出订单样本（用于对比/复算）

```bash
curl -X POST http://localhost:8000/run -d '{
  "script": "workflow005/scripts/export_orders.py",
  "args": ["--start","2026-11-09 00:00:00","--end","2026-11-14 23:59:59",
           "--out","/data/data/workflow005/input","--format","both"]
}'
```

输出：
- `orders_20261109_20261114.json` — 每行一件商品，10 个字段
- `orders_20261109_20261114.xlsx` — 同样内容，给人工核对

## 用样本出参考账单

```bash
# 明细 JSON → 计算表 + PDF
python /scripts/workflow005/scripts/write_from_template.py \
  --detail /data/data/workflow005/input/bill_detail_sample.json \
  --out    /data/data/workflow005/output/60103计算表_样本.xlsx --pdf
```

## 样本对账结果（2026-11-09 ~ 11-14）

| 指标 | 值 |
|---|---|
| 订单 | 198 |
| 明细行 | 277 |
| 可出账 | 179 |
| 阻断 | 19（2 个 SKU 前缀缺失） |
| 国际运费 | 7,336.70 |
| 挂号费 | 4,055.00 |
| 成本 | 4,467.00 |
| 国内运费 | 971.00 |
| 履行价 | 12,774.70 |
| 人民币 | 15,590.09 |
| 账单合计 | 2,331.03 |

参考文件：`output/60103计算表_样本_20261109-20261114.xlsx` / `.pdf`

## 已知口径

- 阻断单（缺配置/缺运价/缺分区）**账单记 0**，不进汇总
- 美元不提前舍入，账单最后 `ROUND(2)`
- 澳洲分区按邮编单条查询（9 个邮编 → 9 次请求，不拉 9091 行）
- 运价区间：下界开、上界闭
