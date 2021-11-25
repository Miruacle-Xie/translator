import excelOperate
import translateTool
import time
import os
import pandas
from urllib import parse

charNumLimit = 6000
separator = '∞'


def transDocument(handle, column, appid, secretKey):
    handle.insertColumn(handle.max_column+1)
    translateTimeStart = time.time()
    cnt = 0
    while cnt < handle.max_row:
        tmp = 0
        string = ""
        for i in range(cnt + 1, handle.max_row + 1):
            if handle.readExcel(i, column) is not None:
                if len(string) + len(handle.readExcel(i, column)) < charNumLimit:
                    string = string + handle.readExcel(i, column) + separator
                    tmp = tmp + 1
                else:
                    tmp = tmp + 1
                    break
            else:
                if len(string) + 1 < charNumLimit:
                    string = string + separator
                    tmp = tmp + 1
                else:
                    tmp = tmp + 1
                    break
        for timeOut in range(100):
            if timeOut != 99:
                translateResult = translateTool.BaiduTranslate('en', 'zh', appid, secretKey)
                Results = translateResult.BdTrans(string)  # 要翻译的词组
                if Results[0]:
                    for i in range(1, len(Results[1].split(separator))):
                        handle.writeExcel(i + cnt, handle.max_column, str(Results[1].split(separator)[i - 1]))
                    break
                else:
                    time.sleep(0.1)
            elif timeOut == 99:
                handle.writeExcel(cnt, handle.max_column, "翻译超时")
                break
        cnt = tmp + cnt
    translateTime = time.time() - translateTimeStart
    print(translateTime)


def readIdPassword(fileName):
    f = open(fileName.replace("\"", ""), "r", encoding='UTF-8')
    string = f.read()
    f.close()
    return string.split("\n")[0][len("APP ID："):], string.split("\n")[1][len("密钥："):]


if __name__ == '__main__':
    if os.path.exists("password.txt"):
        appid, secretKey = readIdPassword("password.txt")
    else:
        flag = input("输入密钥文件路径:自动读取信息, 按回车:手动输入信息\n")
        if flag == "":
            appid = input("APP ID:")
            secretKey = input("密钥:")
        else:
            appid, secretKey = readIdPassword(flag)
    fileName = input("\n需要翻译的文件路径：\n")
    seletColumn = input("\n按回车：全部翻译, 输入数字:指定列数翻译\n")
    try:
        wb = excelOperate.operateExecl(fileName.replace("\"", ""))
        wb.openExcel()
        if seletColumn == "":
            for i in range(1, wb.max_column+1):
                transDocument(wb, i, appid, secretKey)
                wb.saveExcel()
        else:
            for i in range(len(seletColumn.split(" "))):
                wb.openExcel()
                transDocument(wb, int(seletColumn.split(" ")[i]), appid, secretKey)
                wb.saveExcel()
        input("翻译完毕,按任意键结束")
    except Exception as e:
        string = str(e)
        print(string)
        df = pandas.read_excel(fileName.replace("\"", ""), sheet_name=None)
        print(list(df))
        charLen = len("Invalid character ")
        if "Invalid character" in string:
            input("请检查文档sheet名称, 存在非法字符\"{}\", 按任意键结束".format(string[charLen:charLen+1]))
