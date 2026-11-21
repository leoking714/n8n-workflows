# 60103 飞书版（新工程）

**这个目录是独立的**，不 import、也不修改老脚本
`/Users/zuo/n8n/files/automation/scripts/workflow003/generate_table.py`。
老脚本继续用 Excel 跑，互不影响。

---

## 目录结构

```
60103飞书/
├── feishu_client.py        飞书读取层（token / 表 / 分页 / 字段展平）
├── sync_config.py          拉五张配置表 → data/*.json + 体检
├── calc.py                 计算核心（查表、算钱、异常、汇总）
├── writer.py               写 xlsx（逐行公式、商品图、汇总/异常/快照）
├── generate.py             入口：订单 xlsx + data/*.json → 计算表
├── make_local_json.py      开发用：从本地 Excel 造 data/*.json（离线测试）
├── 示例数据/                5 个导入用 Excel（造 JSON 用）
├── 模板/计算表模板.xlsx      写表模板
├── data/                   同步下来的 JSON
├── 输出/                    生成的计算表
└── n8n/                    3 个可导入的工作流
```

## 跑法（离线验证，不连飞书）

```bash
python make_local_json.py      # 用示例数据造 data/*.json
python generate.py             # 取最新订单 → 生成计算表
```

输出：

```
商品图：成功 343 张，失败 0 张
明细 351 行 / 248 单；异常 43 条
汇总 248 单 / 346 件 / 账单 3073.1$
已生成 输出/60103计算表_xxxx.xlsx
```

## 跑法（连飞书）

```bash
export FEISHU_APP_ID=... FEISHU_APP_SECRET=... FEISHU_APP_TOKEN=...
python sync_config.py          # 拉配置 → data/*.json
python generate.py 订单.xlsx    # 出表
```

## 用到的环境变量

| 变量 | 说明 |
|---|---|
| `FEISHU_APP_ID` | 飞书应用 App ID |
| `FEISHU_APP_SECRET` | 飞书应用 App Secret |
| `FEISHU_APP_TOKEN` | 多维表格地址 `/base/` 后面那串（`bascn...`） |
| `TBL_PRODUCTS` 等 | 可选，直接指定 table_id（表名不一致时用） |

表名（必须完全一致，或改用上面的环境变量）：
`商品配置` / `运价表` / `澳洲分区` / `全局参数` / `国家规则`

## 跑法

```bash
cd 60103飞书
export FEISHU_APP_ID=cli_xxx
export FEISHU_APP_SECRET=xxx
export FEISHU_APP_TOKEN=bascnxxx

python sync_config.py --check   # 只体检
python sync_config.py           # 体检 + 落盘到 data/
```

输出示例：

```
拉到：
  商品配置      30 行
  运价表        73 行
  澳洲分区    9091 行
  全局参数      11 行
  国家规则      49 行

✅ 体检通过
已写入 .../data/
```

## 体检会抓这些错

| 现象 | 体检报什么 |
|---|---|
| 表名写错 / 应用没加协作者 | `一条记录都没有` |
| 字段名改过 | `缺少字段 [...]；实际字段 = [...]` |
| SKU 前缀重复 | `SKU前缀 重复 ['CHERRY-DOT']` |
| 运价表下界/上界是文本 | `N 行下界/上界不是数字（字段类型要设成「数字」）` |
| 邮编丢了前导零 | `N 行邮编不是 4 位文本` |

## n8n 工作流

导入 `n8n/飞书配置同步.json` 后有 6 个节点：

```
手动触发 → 配置 → 拉取飞书配置(Execute Command) → 解析结果(Code) → 生成计算表(未启用) → 结束
```

1. 把整个 `60103飞书` 目录拷到 n8n 能访问的位置，默认写的是 `/data/data/workflow003/feishu`
2. 在「配置」节点填 `app_id` / `app_secret` / `app_token`（或改成读环境变量）
3. 执行后看「解析结果」节点的 `counts`

## 下一步（第二阶段）

`generate.py` —— 新工程自己的计算与写表，完全独立实现：

1. 读 `data/*.json`
2. 映射 SKU、查运价、算钱（规则与老脚本一致）
3. 生成计算表 xlsx

写好后把工作流里「生成计算表（待接入）」节点的 `disabled` 去掉即可。
