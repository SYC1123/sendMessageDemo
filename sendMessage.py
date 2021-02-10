# python3.6.5
# 需要引入requests包 ：运行终端->进入python/Scripts ->输入：pip install requests
import time

from ShowapiRequest import ShowapiRequest
import os
from tkinter import *
from tkinter.messagebox import *
import tkinter.filedialog
import easygui as g
import xlrd
import xlwt
import pandas as pd


def xz():
    filename = tkinter.filedialog.askopenfilename()
    if filename != '':
        try:
            book = xlrd.open_workbook(filename)
            sheet = book.sheet_by_index(0)  # 根据顺序获取sheet
            for i in range(1, sheet.nrows):  # 0 1 2 3 4 5
                r = ShowapiRequest("http://route.showapi.com/28-1", "530851", "e75fe3242d6c440d8b64b3a6534da0ae")
                r.addBodyPara("mobile", sheet.row_values(i)[4])
                r.addBodyPara("content", "")
                r.addBodyPara("tNum", "T170317006560")
                r.addBodyPara("big_msg", "")
                res = r.post()
                print(res.text)  # 返回信息
                # print(sheet.row_values(i)[4])
            tkinter.messagebox.showinfo('提示', '提醒完成！')
        except:
            showwarning("警告", "您选择的文件格式错误！")


def xz_before():
    filename = tkinter.filedialog.askopenfilename()
    if filename != '':
        try:
            beforeFileName.set(filename)
        except:
            showwarning("警告", "您选择的文件格式错误！")


def xz_now():
    filename = tkinter.filedialog.askopenfilename()
    if filename != '':
        try:
            nowFileName.set(filename)
        except:
            showwarning("警告", "您选择的文件格式错误！")


def analysis():
    file1 = nowFileName.get()  # 今日
    file2 = beforeFileName.get()  # 昨日
    print(file2)
    if file1 != '' or file2 != '':
        try:
            # weatherfile = os.path.join(os.path.expanduser('~'), "Desktop") + "\\今日名单.xlsx"
            weatherfile = tkinter.filedialog.asksaveasfilename(filetypes=[('xlsx', '*.xlsx')], initialdir='D:\\')
            weatherfile = weatherfile + '.xls'
            writer = pd.ExcelWriter(weatherfile, engine='openpyxl')

            data = pd.DataFrame(pd.read_excel(file1))  # 今天
            data1 = pd.DataFrame(pd.read_excel(file2))  # 昨天
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
            result1 = pd.DataFrame()
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
            else:  # 今天进出黑龙江人数相同
                result2 = pd.DataFrame(columns=[' ', ' '])
                result2.to_excel(writer, sheet_name="人员流动，省内" + str(num) + "人", index=False)
            tkinter.messagebox.showinfo('提示', '分析完成！')
            print(newResult)
            writer.save()
        except:
            showwarning("警告", "您选择的文件格式错误！")
    else:
        showwarning("警告", "请选择昨日和今日的表格！")


def send():
    r = ShowapiRequest("http://route.showapi.com/28-1", "530851", "e75fe3242d6c440d8b64b3a6534da0ae")
    r.addBodyPara("mobile", "18545626763")
    r.addBodyPara("content", "")
    r.addBodyPara("tNum", "T170317006554")
    r.addBodyPara("big_msg", "")
    res = r.post()
    print(res.text)  # 返回信息


if __name__ == '__main__':
    root = Tk()
    root.title("计算机科学技术学院、软件学院今日校园提醒(Code by syc)")
    root.geometry("500x200")
    root.geometry("+500+500")
    beforeFileName = StringVar()
    nowFileName = StringVar()

    fDown = Frame()
    fDown.pack()
    lb = Label(fDown, text='')
    lb.pack()
    btn = Button(fDown, text="未签到提醒", command=xz)
    btn.pack()
    lb = Label(fDown, text='')
    lb.pack()

    fUp = Frame()
    fUp.pack()
    lable1 = Label(fUp, text="昨日表格:", width=8, anchor=E)
    lable1.grid(row=1, column=1)

    lableOne = Label(fUp, textvariable=beforeFileName)
    lableOne.grid(row=1, column=2)

    btn1 = Button(fUp, text="选择", command=xz_before)
    btn1.grid(row=1, column=3)

    fUp1 = Frame()
    fUp1.pack()
    lable2 = Label(fUp1, text="今日表格:", width=8, anchor=E)
    lable2.grid(row=1, column=1)

    lableTwo = Label(fUp1, textvariable=nowFileName)
    lableTwo.grid(row=1, column=2)

    btn2 = Button(fUp1, text="选择", command=xz_now)
    btn2.grid(row=1, column=3)

    fUp2 = Frame()
    fUp2.pack()
    btn3 = Button(fUp2, text="分析", width=10, command=analysis)
    btn3.pack()

    mainloop()
