import openpyxl

FREIGHT = r"./data/60103物流运费-11.2更.xlsx"


def read_excel_demo(file_path):
    # 1. 加载工作簿
    # 注意：data_only=True 表示读取公式计算后的结果。
    # 如果设为 False（默认），则会读取到公式本身（如 "=SUM(A1:A2)"）
    wb = openpyxl.load_workbook(file_path, data_only=True)
    
    # 2. 获取工作表
    # 获取当前活动的工作表（通常是打开时的第一个表）
    sheet = wb.active 
    # 也可以通过名称获取指定工作表： sheet = wb['Sheet1']
    print(f"正在读取工作表: {sheet.title}\n" + "-"*30)
    
    # ================= 常用读取操作 =================
    
    # 操作 A：读取单个单元格
    # 方式1：通过字母+数字坐标 (如 'A1')
    cell_A1 = sheet['A1'].value
    print(f"1. A1 单元格的值为: {cell_A1}")
    
    # 方式2：通过 row 和 column 索引 (注意：行列索引均从 1 开始)
    cell_B2 = sheet.cell(row=2, column=2).value
    print(f"2. B2 (第2行第2列) 的值为: {cell_B2}\n")
    
    
    # 操作 B：遍历读取所有行和列
    print("3. 遍历所有数据 (以元组形式输出每行):")
    # values_only=True 表示直接返回单元格的值，而不是 Cell 对象
    for row in sheet.iter_rows(values_only=True):
        print(row)
    print()
    
    
    # 操作 C：读取指定的单元格范围
    print("4. 读取 A1 到 C3 范围的数据:")
    # 直接通过切片方式获取范围
    for row in sheet['A1':'C3']:
        row_data = [cell.value for cell in row]
        print(row_data)
    print()
    
    
    # 操作 D：获取表格的边界信息
    print(f"5. 表格信息 -> 最大行数: {sheet.max_row}, 最大列数: {sheet.max_column}")
    
    # 关闭工作簿（释放资源）
    wb.close()

# ================= 运行测试 =================
if __name__ == "__main__":

    read_excel_demo(FREIGHT)
