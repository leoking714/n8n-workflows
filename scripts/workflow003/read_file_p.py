import pandas as pd
import numpy as np

# ================= 1. 全局显示设置 =================
# 设置显示所有行和列，防止在控制台打印时被截断显示为 '...'
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
# 设置显示完整的单元格内容，防止长文本（如备注）被截断
pd.set_option('display.max_colwidth', None) 

# ================= 2. 读取数据 =================
# 提示：请确保脚本运行目录下存在该文件，或将其修改为您的绝对路径 (如 r"C:\Users\...\60103物流运费-11.2更.xlsx")
file_path = r"./data/60103物流运费-11.2更.xlsx" 

# 【性能优化】：在读取时直接限制只解析前13列 (usecols=range(13))
# 这比先全部读取再用 iloc 截取更省内存、速度更快，且能自动忽略后面无用的空白列
print("⏳ 正在读取数据...")
df = pd.read_excel(file_path, header=0, usecols=range(13))

print(f"📊 原始数据形状: {df.shape} (行数, 列数)")
print("原始数据前3行预览:\n", df.head(3), "\n")

# ================= 3. 核心数据清洗步骤 =================

# 步骤 3.1: 处理第一列的合并单元格 (向下填充)
# 从知识库可见，如“挪威”、“比利时”等国家的部分行第一列为空，ffill() 会用正上方的国家名填充下方的空值
first_col_name = df.columns[0] # 动态获取第一列的真实列名 (即 '60103国家报价')
df[first_col_name] = df[first_col_name].ffill()

# 步骤 3.2: 处理“伪空值”
# 物流数据中常混入字符串形式的空值，将它们统一替换为 Pandas 认识的 NaN，以便后续统一清理
df = df.replace(['None', 'null', 'NA', '', ' '], np.nan)

# 步骤 3.3: 删除全空行
# how='all' 表示只有当这一行的所有13列都是空值(NaN)时，才删除该行
# 这能有效清理掉文件中那些纯粹的空白占位行
df = df.dropna(how='all')

# 步骤 3.4: 去除文本列首尾的不可见空格 (强烈推荐)
# 避免因为 "瑞典 " 和 "瑞典" 导致后续分组或匹配失败
str_cols = df.select_dtypes(include=['object', 'string']).columns
df[str_cols] = df[str_cols].apply(lambda x: x.str.strip() if x.dtype == 'object' else x)

# 步骤 3.5: 重置索引
# 删除行后，原来的行索引会断开（比如变成 0, 1, 3, 4），重新生成连续的 0, 1, 2...
df = df.reset_index(drop=True)

# ================= 4. 输出与保存 =================
print("="*60)
print(f"✅ 清洗后数据形状: {df.shape}")
print("清洗后的完整数据预览:\n", df.head(10)) # 预览前10行

# 将清洗后的数据保存为新的 Excel 文件（不保存行索引，使用 openpyxl 引擎）
output_path = r"./60103物流运费-11.2更_已清洗.xlsx"
df.to_excel(output_path, index=False, engine='openpyxl')
print(f"\n💾 数据已成功清洗并保存至: {output_path}")