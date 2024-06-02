import openpyxl
import re
pattern_normal = r'^(\d+)-(\d+)K$'
pattern_price= r'^(\d+)-(\d+)K·(\d+)薪$'
pattern_day = r'^(\d+)-(\d+)元/天$'
pattern_week = r'^(\d+)-(\d+)元/周$'
pattern_month = r'^(\d+)-(\d+)元/月$'
pattern_hour = r'^(\d+)-(\d+)元/时$'
wb=openpyxl.load_workbook('boss.xlsx')
sheet=wb['react']
# 获取F列的每个数据
low_price=0
high_price=0
total_number=0
for cell in sheet['F']:
    if cell.value is not None:
        total_number+=1
        match = re.match(pattern_normal, str(cell.value))
        if match:
            num1 = match.group(1)  # 获取第一个捕获组中的内容（即前面的数字）
            num2 = match.group(2)  # 获取第二个捕获组中的内容（即后面的数字）
            low_price+=(float(num1)/300)*12
            high_price+=(float(num1)/300)*12
            print(low_price,high_price)
            continue
        match = re.match(pattern_price, str(cell.value))
        if match:
            num1 = match.group(1)  # 获取第一个捕获组中的内容（即前面的数字）
            num2 = match.group(2)  # 获取第二个捕获组中的内容（即后面的数字）
            num3 = match.group(3)  # 获取第三个捕获组中的内容（即薪水数字）
            low_price+=float(int(num1)*int(num3))/300
            high_price+=float(int(num2)*int(num3))/300
            print(low_price, high_price)
            continue
        match = re.match(pattern_day, str(cell.value))
        if match:
            num1 = match.group(1)  # 获取第一个捕获组中的内容（即前面的数字）
            num2 = match.group(2)  # 获取第二个捕获组中的内容（即后面的数字）
            low_price+=(float(float(num1)/300)*365)/1000
            high_price+=(float(float(num1)/300)*365)/1000
            print(low_price, high_price)
            continue
        match = re.match(pattern_week, str(cell.value))
        if match:
            num1 = match.group(1)  # 获取第一个捕获组中的内容（即前面的数字）
            num2 = match.group(2)  # 获取第二个捕获组中的内容（即后面的数字）
            low_price += (float(float(num1)/300) * 48)/1000
            high_price += (float(float(num1)/300) * 48)/1000
            print(low_price, high_price)
            continue
        match = re.match(pattern_month, str(cell.value))
        if match:
            num1 = match.group(1)  # 获取第一个捕获组中的内容（即前面的数字）
            num2 = match.group(2)  # 获取第二个捕获组中的内容（即后面的数字）
            low_price += (float(float(num1)/300)*12)/1000
            high_price += (float(float(num1)/300)*12)/1000
            print(low_price, high_price)
            continue
        match = re.match(pattern_hour, str(cell.value))
        if match:
            num1 = match.group(1)  # 获取第一个捕获组中的内容（即前面的数字）
            num2 = match.group(2)  # 获取第二个捕获组中的内容（即后面的数字）
            low_price += (float(float(num1)/1000)*3336)/300
            high_price += (float(float(num1)/1000)*3336)/300
            print(low_price, high_price)
            continue
        print("都匹配不上,有错误")
print("最终平均薪资:")
print(low_price/12,high_price/12)