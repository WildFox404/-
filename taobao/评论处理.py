#h5/mtop.alibaba.review.list.for.new.pc.detail/1.0/
#h5/mtop.alibaba.review.list.for.new.pc.detail/1.0/
from DrissionPage import WebPage
import openpyxl
import time
#https://detail.tmall.com/item.htm?id=764262180203

def login(page):
    page.get("https://login.taobao.com/member/login.jhtml")
    time.sleep(10)
    page.ele("@class=search-combobox-input", timeout=20).input("abc")




if __name__ == '__main__':
    column_list=['I',"J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ"]
    wb = openpyxl.load_workbook('taobao_test.xlsx')
    sheet = wb['Sheet1']
    page = WebPage()
    login(page)
    for cell in sheet['A']:
        if cell.row>=959:
            texts_len = 0
            time.sleep(1)
            url = cell.value
            page.get(url)
            time.sleep(1)
            try:
                page.ele('xpath=//*[@id="root"]/div/div[2]/div[2]/div[2]/div[1]/div/div/div[2]/span').click()
            except:
                continue
            texts=page.eles("@class=Comment--content--15w7fKj")
            # 遍历texts列表，并逐个将值写入到指定的列
            for i in texts:
                texts_len +=1
            print(texts_len)
            for link in texts:
                texts_len-=1
                if texts_len >= 0:
                    # 获取目标列，并将link.text的值写入
                    sheet[column_list[texts_len-1] + str(cell.row)] = str(link.text)
            print()
            wb.save('taobao_test.xlsx')