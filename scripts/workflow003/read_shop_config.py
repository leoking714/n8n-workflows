import pandas as pd

# 设置显示所有行
pd.set_option('display.max_rows', None)

# 设置显示所有列
pd.set_option('display.max_columns', None)

df = pd.read_excel(r"./data/商品配置-11.2更.xlsx", header=0)
df = df.iloc[:,2:]

# 步骤 A: 提取第 0 行和第 1 行，并将第 0 行的 NaN 替换为空字符串 ''
row0 = df.iloc[0].fillna('') 
row1 = df.iloc[1].astype(str)

# 步骤 B: 使用列表推导式拼接列名
# 逻辑：如果 row0 的值为空，则只保留 row1 的值；否则拼接为 "row1【row0】"
new_columns = [
    f"{v1}【{v0}】" if v0 != '' else v1 
    for v0, v1 in zip(row0, row1)
]

# 步骤 C: 将新列名赋值给 DataFrame
df.columns = new_columns

# 步骤 D: 删除原来的前两行（它们已经变成表头了），并重置索引
df = df.iloc[2:].reset_index(drop=True)


print(df)