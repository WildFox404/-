from DrissionPage import WebPage
import openpyxl
import time
class Boss():
    def __init__(self,workbook,keyword):
        self.page = WebPage()
        self.sheet=workbook[f"{keyword}"]
        self.keyword=keyword

    def login_in(self):
        self.page.get('https://www.zhipin.com/web/user/?ka=header-login')
        try:
            self.page.ele("@class=btn-sign-switch ewm-switch").click()
        except:
            print("可能已经登录->直接跳转")
        self.page.get(f"https://www.zhipin.com/web/geek/job?query={self.keyword}&city=100010000")
    def get_data(self):
        data=self.page.listen.wait()
        if data.response.body==None:
            print("未获取到数据")
            self.get_data()
        else:
            return data.response.body
    def put_data(self,data):
        job_list = data['zpData']['jobList']
        for i in range(len(job_list)):
            row_list=[]
            print(job_list[i]['bossCert'], end="\t")
            row_list.append(job_list[i]['bossCert'])
            print(job_list[i]['bossName'], end="\t")
            row_list.append(job_list[i]['bossName'])
            print(job_list[i]['bossTitle'], end="\t")
            row_list.append(job_list[i]['bossTitle'])
            print(job_list[i]['goldHunter'], end="\t")
            row_list.append(job_list[i]['goldHunter'])
            print(job_list[i]['jobName'], end="\t")
            row_list.append(job_list[i]['jobName'])
            print(job_list[i]['salaryDesc'], end="\t")
            row_list.append(job_list[i]['salaryDesc'])
            print(job_list[i]['jobLabels'], end="\t")
            row_list.append(job_list[i]['jobLabels'])
            print(job_list[i]['iconWord'], end="\t")
            row_list.append(job_list[i]['iconWord'])
            print(job_list[i]['skills'], end="\t")
            row_list.append(job_list[i]['skills'])
            print(job_list[i]['jobExperience'], end="\t")
            row_list.append(job_list[i]['jobExperience'])
            print(job_list[i]['daysPerWeekDesc'], end="\t")
            row_list.append(job_list[i]['daysPerWeekDesc'])
            print(job_list[i]['leastMonthDesc'], end="\t")
            row_list.append(job_list[i]['leastMonthDesc'])
            print(job_list[i]['jobDegree'], end="\t")
            row_list.append(job_list[i]['jobDegree'])
            print(job_list[i]['cityName'], end="\t")
            row_list.append(job_list[i]['cityName'])
            print(job_list[i]['areaDistrict'], end="\t")
            row_list.append(job_list[i]['areaDistrict'])
            print(job_list[i]['businessDistrict'], end="\t")
            row_list.append(job_list[i]['businessDistrict'])
            print(job_list[i]['brandName'], end="\t")
            row_list.append(job_list[i]['brandName'])
            print(job_list[i]['brandStageName'], end="\t")
            row_list.append(job_list[i]['brandStageName'])
            print(job_list[i]['brandIndustry'], end="\t")
            row_list.append(job_list[i]['brandIndustry'])
            print(job_list[i]['brandScaleName'], end="\t")
            row_list.append(job_list[i]['brandScaleName'])
            print(job_list[i]['welfareList'])
            row_list.append(job_list[i]['welfareList'])
            row_list=[str(item) for item in row_list]
            self.sheet.append(row_list)


if __name__ == '__main__':
    wb = openpyxl.load_workbook("boss.xlsx")
    keyword="android"
    # 创建一个新的 sheet
    new_sheet = wb.create_sheet(title=f'{keyword}')
    boss = Boss(wb,keyword)
    boss.page.listen.start("zpgeek/search/joblist.json")
    boss.page.get(f"https://www.zhipin.com/web/geek/job?query={keyword}&city=100010000")
    #必要时进行登录
    boss.login_in()
    time.sleep(6)
    for i in range(10):
        data_json=boss.get_data()
        if data_json==None:
            print("数据为空,获取失败")
        print(f"第{i}页数据获取成功")
        boss.page.ele("@class=ui-icon-arrow-right").click()
        boss.put_data(data_json)
        print(f"第{i}页数据已保存")
    wb.save("boss.xlsx")