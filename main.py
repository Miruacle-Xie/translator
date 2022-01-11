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
        print("当前翻译第{}行".format(cnt))
        tmp = 0
        string = ""
        cellValue = ""
        for i in range(cnt + 1, handle.max_row + 1):
            # print("handle.readExcel(i, column):{}, i:{}".format(handle.readExcel(i, column), i))
            if handle.readExcel(i, column) is not None:
                if isinstance(handle.readExcel(i, column), int):
                    cellValue = str(handle.readExcel(i, column))
                else:
                    cellValue = handle.readExcel(i, column)
                if len(string) + len(cellValue) < charNumLimit:
                    string = string + cellValue + separator
                    tmp = tmp + 1
                else:
                    tmp = tmp
                    break
            else:
                if len(string) + 1 < charNumLimit:
                    string = string + separator
                    tmp = tmp + 1
                else:
                    tmp = tmp
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
        print("共{}行".format(wb.max_row))
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
        wb.saveExcel()
        string = str(e)
        print(string)
        df = pandas.read_excel(fileName.replace("\"", ""), sheet_name=None)
        print(list(df))
        charLen = len("Invalid character ")
        if "Invalid character" in string:
            input("请检查文档sheet名称, 存在非法字符\"{}\", 按任意键结束".format(string[charLen:charLen+1]))
        else:
            input("sorry, 出bug了, QAQ")
