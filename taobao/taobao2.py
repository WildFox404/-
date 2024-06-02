import time
import json
import openpyxl
from DrissionPage import WebPage
import re
class TaoBao():
    def __init__(self,workbook):
        self.page=WebPage()
        self.sheet=workbook["Sheet1"]
    def login(self):
        self.page.get("https://login.taobao.com/member/login.jhtml")
        time.sleep(10)
        self.page.ele("@class=search-combobox-input",timeout=20).input("abc")

    def get_data(self):
        data=self.page.listen.wait(timeout=2)
        if data==False:
        # if data.response.body==None:
            print("未获取到数据")
            return""
        else:
            return data.response.body
    def search_data(self,data):
        try:
            result_list=[]
            # 查找JSON数据起始位置
            start_index = data.find("{")
            # 查找JSON数据终止位置
            end_index = data.rfind("}") + 1

            # 如果找到起始和终止位置，则提取JSON数据
            if start_index != -1 and end_index != -1:
                json_data = data[start_index:end_index]
            else:
                print("未找到JSON数据")
            json_data = json_data.encode('utf-8')
            json_data = json.loads(json_data)
            data_list = json_data['data']['data']['items']

            for i in range(len(data_list)):
                sell_fuzzy_num = int(re.search(r'\d+', data_list[i]['material']['sellFuzzy']).group())
                if sell_fuzzy_num < 5:
                    print(data_list[i]['material']['eurl'])  # 店铺链接
                    print(data_list[i]['material']['sellFuzzy'])  # 销售额
                    result_list.append(data_list[i]['material']['eurl'])
            return result_list
        except:
            print("json数据出错")
            time.sleep(5)
            return []

def get_data(page):
    data=page.listen.wait(timeout=3)
    if data.response.body==None:
        return False
    else:
        return True
def get_shop_url(data):
    # 使用正则表达式匹配
    pattern = r"'pcShopUrl':\s*'(.*?)'"
    result = re.search(pattern, data)

    # 提取匹配到的内容
    if result:
        extracted_content = result.group(1)
        return extracted_content
    else:
        return ""
def get_data_result(page):
    data = page.listen.wait(timeout=2)
    if data.response.body == None:
        print("未获取到数据")
    else:
        return data.response.body
if __name__ == '__main__':
    wb = openpyxl.load_workbook("taobao2.xlsx")
    taobao=TaoBao(wb)
    taobao.login()
    taobao.page.listen.start("h5/mtop.alimama.abyss.unionpage.get/1.0")
    page = WebPage()
    page.listen.start(targets="h5/mtop.taobao.pcdetail.data.get/1.0/")
    # shopPage=WebPage()
    # shopPage.listen.start(targets="https://gw.alicdn.com/tps/TB1hbG6PpXXXXauapXXXXXXXXXX-88-24.png")
    for i in range(100):
        time.sleep(1)
        taobao.page.get(f"https://uland.taobao.com/sem/tbsearch?keyword=%E6%89%8B%E6%9C%BA%E5%BE%AE%E4%B9%90%E5%B0%8F%E7%A8%8B%E5%BA%8F%E5%AE%B6%E4%B9%A1%E5%B9%BF%E4%B8%9C%E9%BA%BB%E5%B0%86&pnum={i+1}")
        data=taobao.get_data()
        if data=="":
            time.sleep(2)
            continue
        print("第{}页数据获取成功\n".format(i+1))
        result_list = taobao.search_data(data)
        for eurl in result_list:
            sheet_result=[]
            try:
                page.get(eurl)
                #https://gw.alicdn.com/tps/TB1hbG6PpXXXXauapXXXXXXXXXX-88-24.png
                print("开始检测是否是新店")
                page.ele('xpath=/html/body/div[3]/div/div[2]/div[1]/a/div[1]/div[1]/div').click()
                shop_data=get_data_result(page)
                shop_url=get_shop_url(shop_data)
                if shop_url != "":
                #     shopPage.get(shop_url)
                    # if get_data(shopPage):
                    #     print("该店铺为新店")
                    sheet_result.append("是淘宝店")
                    sheet_result.append(shop_url)
                    taobao.sheet.append(sheet_result)
            except:
                print("该店铺未获取到内容")
                sheet_result.append("不是淘宝店")
                sheet_result.append(eurl)
                taobao.sheet.append(sheet_result)
        print("第{}页数据写入成功\n".format(i+1))
        wb.save("taobao2.xlsx")