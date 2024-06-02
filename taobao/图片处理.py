from openpyxl import *
from openpyxl.drawing.image import Image
import requests
wb=load_workbook('taobao_test.xlsx')
sheet=wb['Sheet1']
for cell in sheet['A']:
    if cell.value and not cell.value.startswith('http'):
        cell.value = 'https:' + cell.value
wb.save('taobao_test.xlsx')
print("ok")