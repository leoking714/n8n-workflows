import pandas as pd

# 1. 显示所有行
pd.set_option('display.max_rows', None)

# 2. 显示所有列
pd.set_option('display.max_columns', None)

df = pd.read_excel("/data/data/10.27账单61287.xlsx", header=None)
first_col = df.iloc[:, 0]
# print(first_col)
target_rows = first_col[first_col.astype(str).str.strip() == 'No.'].index
target_rows_ids = target_rows[0]
df = pd.read_excel("/data/data/10.27账单61287.xlsx", header=target_rows_ids)
# print(df)
# 清理尾部数据
nan_mask = df["Item Name"].isna()
if nan_mask.any():
    first_nan_pos = nan_mask.idxmax()
    df = df.iloc[:first_nan_pos]

# print(df)
df = df.ffill()
df['Order Number'] = df['Order Number'].astype('Int64')
df['Transaction Number'] = df['Transaction Number'].astype('Int64')

json_str = df.to_json(orient='records', force_ascii=False, indent=4)
print(json_str)
