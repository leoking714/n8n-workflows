# 文件备份

最后备份：2026-11-24 11:20:57

## 目录

| 目录 | 内容 |
|---|---|
| `workflows/` | n8n 已发布工作流 JSON（敏感字段已脱敏）|
| `scripts/` | 自动化 Python 脚本 |
| `templates/` | 计算表模板 / Bill 模板 / 页头图 |
| `config/` | 全局参数快照、飞书 Base 配置（脱敏）|
| `logs/` | 最近 10 次账单明细 JSON |

## 已发布工作流

| 名称 | 节点数 | 文件 |
|---|---|---|
| R1 | 21 | `workflows/R1-RnDESn3R.json` |
| airtable写入采购价和运费-workflow001 | 27 | `workflows/airtable写入采购价和运费-workflow001-NjNTrPqe.json` |
| workflow005 · v7版（抓包） | 54 | `workflows/workflow005-v7版-抓包-1vtebK94.json` |
| 插件更新马帮cookie | 3 | `workflows/插件更新马帮cookie-K2fxcraF.json` |
| 更新马帮cookie | 3 | `workflows/更新马帮cookie-rUWp1Guv.json` |
| 模板｜全局错误捕获 | 10 | `workflows/模板-全局错误捕获-HREvfdug.json` |
| 生成60103客户对账单 | 28 | `workflows/生成60103客户对账单-25miDNyx.json` |
| 飞书长链接测试 | 2 | `workflows/飞书长链接测试-p6xIfVmN.json` |
| workflow005 · v7版（新表/运价按重量段） | 44 | `workflows/workflow005-v7版-新表-运价按重量段-c7cb10db.json` |
| workflow007 · 核对后生成Bill PDF | 7 | `workflows/workflow007-核对后生成Bill-PDF-eb3fb6d6.json` |

## 恢复

```bash
# 工作流
docker exec n8n-n8n-1 n8n import:workflow --input=/data/.../workflows/xxx.json
# 脚本（挂载目录直接覆盖）
cp -R scripts/workflow005 /Users/zuo/n8n/files/automation/scripts/
# 模板
cp templates/计算表模板.xlsx /Users/zuo/n8n/files/data/workflow005/config/
```

> 密钥类字段在导出时替换为 `***REDACTED***`，恢复后需重新填写。
