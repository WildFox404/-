import openpyxl

# 打开 Excel 文件
workbook = openpyxl.load_workbook('taobao3.xlsx')
sheet = workbook.active

# 创建一个集合用于存储唯一值
unique_values = set()

# 遍历第一列的每个单元格，并将唯一值添加到集合中
for cell in sheet['A']:
    if cell.value:
        unique_values.add(cell.value)

# 清空第一列的数据
for cell in sheet['A']:
    cell.value = None

# 将唯一值重新放入第一列
row = 1
for value in unique_values:
    sheet.cell(row, 1, value)
    row += 1

# 保存修改后的 Excel 文件
workbook.save('taobao3.xlsx')