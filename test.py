import openpyxl
import pandas as pd
import os

try:
    weatherfile = os.path.join(os.path.expanduser('~'), "Desktop") + "\\名单.xlsx"
    writer = pd.ExcelWriter(weatherfile, engine='openpyxl')

    data = pd.DataFrame(pd.read_excel('C:/Users/syc87/Desktop/新建文件夹/【2月1日】疫情防控健康信息统计 (1).xlsx'))  # 今天
    data1 = pd.DataFrame(pd.read_excel('C:/Users/syc87/Desktop/新建文件夹/【1月31日】疫情防控健康信息统计.xlsx'))  # 昨天
    # # 直接筛选方法
    some = data[(data['健康状况'] == '新冠肺炎疑似') | (data['健康状况'] == '新冠肺炎确诊')]
    newResult = some[['学号', '姓名', '手机号', '班级', '健康状况']]
    newResult.to_excel(writer, sheet_name="健康状况", index=False)

    some = data[(data['密接情况'] == '与确诊或无症状病例密接')]
    newResult = some[['学号', '姓名', '手机号', '班级', '密接情况']]
    newResult.to_excel(writer, sheet_name="密接情况", index=False)

    some = data[(data['是否发热/呼吸困难/乏力'] == '是')]
    newResult = some[['学号', '姓名', '手机号', '班级', '是否发热/呼吸困难/乏力']]
    newResult.to_excel(writer, sheet_name="是否发热或呼吸困难或乏力", index=False)

    some = data[(data['是否被隔离'] == '是')]
    newResult = some[['学号', '姓名', '手机号', '班级', '是否被隔离']]
    newResult.to_excel(writer, sheet_name="是否被隔离", index=False)

    some = data[(data['目前所在地是否为中高风险地区'] == '是')]
    newResult = some[['学号', '姓名', '手机号', '班级', '目前所在地是否为中高风险地区', '目前所在地区']]
    newResult.to_excel(writer, sheet_name="目前所在地是否为中高风险地区", index=False)

    some = data[(data['定位信息'].str.contains('黑龙江'))]  # 今日
    num = some.shape[0]
    print("今日" + str(num))

    some1 = data1[(data1['定位信息'].str.contains('黑龙江'))]  # 昨日
    num1 = some1.shape[0]
    print("昨日" + str(num1))
    result1 = None
    if num > num1:  # 今天有人进来
        result = pd.merge(some, some1.loc[:, ['学号', '定位信息']], how='left', on='学号')
        newResult = result[result.isnull().T.any()]
        result1 = pd.merge(newResult, data1.loc[:, ['学号', '定位信息']], how='left', on='学号')
        result2 = result1[['学号', '姓名', '手机号', '班级', '定位信息', '定位信息_x']]
        result2.to_excel(writer, sheet_name="人员流动，省内" + str(num) + "人", index=False)
    elif num < num1:  # 今天有人离开
        result = pd.merge(some1, some.loc[:, ['学号', '定位信息']], how='left', on='学号')
        newResult = result[result.isnull().T.any()]
        result1 = pd.merge(newResult, data.loc[:, ['学号', '定位信息']], how='left', on='学号')
        result2 = result1[['学号', '姓名', '手机号', '班级', '定位信息_x', '定位信息']]
        result2.to_excel(writer, sheet_name="人员流动，省内" + str(num) + "人", index=False)
    # else:  # 今天进出黑龙江人数相同
    # result = pd.merge(some, some1.loc[:, ['学号', '定位信息']], how='left', on='学号')
    # newResult = result[result.isnull().T.any()]
    # result1 = pd.merge(newResult, data1.loc[:, ['学号', '定位信息']], how='left', on='学号')

    print(newResult)
    writer.save()
except:
    print(123)
