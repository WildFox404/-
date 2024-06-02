import time
import json
import openpyxl
from DrissionPage import WebPage

class TaoBao():
    def __init__(self,workbook):
        self.page=WebPage()
        self.sheet=workbook["Sheet1"]
    def login(self):
        self.page.get("https://login.taobao.com/member/login.jhtml")
        time.sleep(10)
        self.page.ele("@class=search-combobox-input",timeout=20).input("abc")

    def get_data(self):
        data=self.page.listen.wait()
        if data.response.body==None:
            print("未获取到数据")
            self.get_data()
        else:
            return data.response.body
    def put_data(self,data):
        try:
            # 查找JSON数据起始位置
            start_index = data.find("{")
            # 查找JSON数据终止位置
            end_index = data.rfind("}") + 1
            # 如果找到起始和终止位置，则提取JSON数据
            if start_index != -1 and end_index != -1:
                json_data = data[start_index:end_index]
            else:
                print("未找到JSON数据")
            # json_data = json_data.encode('utf-8')
            json_data = json.loads(json_data)
            print(json_data)
            data_list = json_data['data']['itemsArray']
        except:
            print("json数据出错")
            time.sleep(5)
            return
        for i in range(len(data_list)):
            try:
                temp_list=[]
                temp_list.append(data_list[i]['auctionURL'])
                temp_list.append(data_list[i]['shopInfo']['title'])
                temp_list.append(data_list[i]['shopInfo']['url'])
                temp_list.append(data_list[i]['title'])
                temp_list.append(data_list[i]['price'])
                temp_list.append(data_list[i]['procity'])
                temp_list.append(data_list[i]['realSales'])
                if 'pic_path' in data_list[i]:
                    temp_list.append(data_list[i]['pic_path'])
                self.sheet.append(temp_list)
            except Exception as e:
                print(e)
                continue

if __name__ == '__main__':
    wb = openpyxl.load_workbook("taobao_test.xlsx")
    taobao=TaoBao(wb)
    taobao.login()
    taobao.page.listen.start("h5/mtop.relationrecommend.wirelessrecommend.recommend/2.0")
    for i in range(100):
        time.sleep(1)
        taobao.page.get(f"https://s.taobao.com/search?page={i+1}&q=%E6%88%92%E7%83%9F&tab=all")
        data=taobao.get_data()
        print("第{}页数据获取成功\n".format(i+1))
        taobao.put_data(data)
        print("第{}页数据写入成功\n".format(i+1))
        wb.save("taobao_test.xlsx")
