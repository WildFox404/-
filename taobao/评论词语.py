import openpyxl
import jieba
from collections import Counter

def  data_process(data,dict_result):
    # 使用jieba进行中文分词
    words = jieba.lcut(data)
    # 过滤掉单个字的词语
    words = [word for word in words if len(word) > 1]
    # 使用Counter统计词语出现的频率
    word_freq = Counter(words)
    # 提取出现频率最高的前N个词
    top_n = 10
    top_words = word_freq.most_common(top_n)
    for word, freq in top_words:
        if word in dict_result:
            # 如果word已经存在于字典中，累加freq
            dict_result[word] += freq
        else:
            # 如果word不存在于字典中，添加word并设置freq
            dict_result[word] = freq
if __name__ == '__main__':
    dict_result={}
    wb=openpyxl.load_workbook('taobao_test.xlsx')
    sheet=wb['Sheet1']
    column_list=['I',"J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ"]
    for i in range(1,sheet.max_row+1):
        for j in range(1,len(column_list)):
            if sheet[column_list[j] + str(i)].value==None:
                print(f"第{i}行第{j}个数据为空")
            else:
                data=sheet[column_list[j] + str(i)].value
                data_process(data,dict_result)
        print(f"第{i}行数据处理完成")
    dict_result = dict(sorted(dict_result.items(), key=lambda x: x[1], reverse=True))
    top = dict(list(dict_result.items())[:100])
    print(dict_result)
    print(top)
    # 将字典按值降序排序
    sorted_data = sorted(top.items(), key=lambda x: x[1], reverse=True)

    # 将排序后的数据输出为所需格式
    for item in sorted_data:
        print(f'("{item[0]}", {item[1]}),')