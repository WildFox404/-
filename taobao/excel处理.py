from openpyxl import *
from openpyxl.drawing.image import Image
import requests
wb=load_workbook('taobao_test.xlsx')
sheet=wb['Sheet1']
for cell in sheet['D']:
    if cell.value is not None and "<span class=H>戒烟</span>" in cell.value:
        cell.value = cell.value.replace("<span class=H>戒烟</span>", "  戒烟  ")
# for cell in sheet['H']:
#     if cell.value is not None:
#         # 发送HTTP请求获取图片数据
#         response = requests.get(cell.value)
#         if response.status_code == 200:
#             # 保存图片到本地
#             with open('image.png', 'wb') as file:
#                 file.write(response.content)
#             print("图片下载成功")
#             # 插入图片
#             img = Image('image.png')  # 指定图片路径
#             sheet.add_image(img, 'I' + str(cell.row))
wb.save('taobao_test.xlsx')
print("ok")
