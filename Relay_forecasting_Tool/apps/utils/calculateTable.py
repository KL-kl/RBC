import os
# import platform
# import pythoncom
# import win32com.client as win32
# from win32com.client import Dispatch


# 把数字转换成相应的字符,1-->'A' 27-->'AA'
def changeNumToChar2(toSmallChar=None, toBigChar=None):
    init_number = 0
    increment = 0
    res_char = ''
    char1 = ''
    if not toSmallChar and not toBigChar:
        return ''
    else:
        if toSmallChar:
            init_number = toSmallChar
            increment = ord('A') - 1
        else:
            init_number = toBigChar
            increment = ord('a') - 1

    shang, yu = divmod(init_number, 26)

    if yu > 0:
        char = chr(yu + increment)
    else:
        char = chr(26 + increment)

    if shang > 1 and yu == 0:
        char1 = chr(shang + increment - 1)
    elif shang >= 1 and yu > 0:
        char1 = chr(shang + increment)

    if char1 != '':
        res_char = char1 + char
    else:
        res_char = char

    return res_char


# sheet 相应单元格取最大
def getMaxFiro(l=[]):
    res_char = 'MAX('

    for i in l:
        index = l.index(i)
        if index != (len(l) - 1):
            res_char = res_char + i + '!@,'
        else:
            res_char = res_char + i + '!@)'

    return res_char


# 多张sheet页相应位置求和
def getSumFiro(l=[]):

    res_char = 'SUM('

    remove_list = ['F', '点到点流量', 'Mysheet']

    for i in l:
        if i in remove_list:
            l.remove(i)
        else:
            index = l.index(i)
            if index != (len(l) - 1):
                res_char = res_char + i + '!@,'
            else:
                res_char = res_char + i + '!@)'

    return res_char


# 替换公式内的位置
def getFiro2(fro, i, j):
    liebiao = changeNumToChar2(i + 1) + str(j + 1)
    res_char = fro.replace('@', liebiao)


    return res_char


# 计算点到点流量
def getFlow(i, j, num,sheetname):
    row = str(i + 1)
    col = changeNumToChar2(j + 1)
    len_char = changeNumToChar2(num)
    res_char = 'IF(LEFT($A' + row + ',IF(ISNUMBER(FIND("1",$A' + row + ')),FIND("1",$A' + row + '),FIND("2",$A' + row + '))-1)=LEFT(' + col + '$1,IF(ISNUMBER(FIND("1",' + col + '$1)),FIND("1",' + col + '$1),FIND("2",' + col + '$1))-1),0,IF(MID($A' + row + ',IF(ISNUMBER(FIND("1",$A' + row + ')),FIND("1",$A' + row + '),FIND("2",$A' + row + ')),1)=MID(' + col + '$1,IF(ISNUMBER(FIND("1",' + col + '$1)),FIND("1",' + col + '$1),FIND("2",' + col + '$1)),1),OFFSET('+sheetname+'!$A$1,MATCH(VLOOKUP(LEFT($A' + row + ',IF(ISNUMBER(FIND("1",$A' + row + ')),FIND("1",$A' + row + '),FIND("2",$A' + row + '))-1),F!$D$2:$E$32,2,FALSE),'+sheetname+'!$A$2:$A$' + str(
        num) + ',0),MATCH(VLOOKUP(LEFT(' + col + '$1,IF(ISNUMBER(FIND("1",' + col + '$1)),FIND("1",' + col + '$1),FIND("2",' + col + '$1))-1),F!$D$2:$E$32,2,FALSE),'+sheetname+'!$B$1:$' + len_char + '$1,0))/2,0))'
    # res_char = 'SUM(流量需求汇总_L2!$B$3)'
    # print(res_char)
    return res_char


# xls转换xlsx
# def xlstoxlsx(fname):
#     pythoncom.CoInitialize()
#     # 判断当前系统
#     os_name = platform.system()
#     if os_name == 'Windows':
#
#         report = os.path.join(os.getcwd(), fname)  # 获取当前路径,xxx换成你的文件名
#
#         excel = win32.gencache.EnsureDispatch('Excel.Application')
#
#         excel.Visible = 0
#         excel.DisplayAlerts = 0
#
#         wb = excel.Workbooks.Open(report)
#
#         wb.SaveAs(report + 'x', FileFormat=51)
#
#         wb.Close()
#
#         excel.Application.Quit()
#
#     elif os_name == 'Linux':
#
#         report = os.path.join(os.getcwd(), fname).replace('\\', '/')
#
#         excel = win32.gencache.EnsureDispatch('Excel.Application')
#         excel.Visible = 0
#         excel.DisplayAlerts = 0
#
#         wb = excel.Workbooks.Open(report)
#
#         wb.SaveAs(report + 'x', FileFormat=51)
#
#         wb.Close()
#
#         excel.Application.Quit()
#     # 释放资源
#     pythoncom.CoUninitialize()

#xlsx转成xls
# def xlsxtoxls(fname):
#     pythoncom.CoInitialize()
#     # 判断当前系统
#     os_name = platform.system()
#     if os_name == 'Windows':
#
#         report = os.path.join(os.getcwd(), fname)  # 获取当前路径,xxx换成你的文件名
#
#         excel = win32.gencache.EnsureDispatch('Excel.Application')
#
#         excel.Visible = 0
#         excel.DisplayAlerts = 0
#
#         wb = excel.Workbooks.Open(report)
#
#         wb.SaveAs(report[:-1], FileFormat=56)
#
#         wb.Close()
#
#         excel.Application.Quit()
#
#     elif os_name == 'Linux':
#
#         report = os.path.join(os.getcwd(), fname).replace('\\', '/')
#
#         excel = win32.gencache.EnsureDispatch('Excel.Application')
#         excel.Visible = 0
#         excel.DisplayAlerts = 0
#
#         wb = excel.Workbooks.Open(report)
#
#         wb.SaveAs(report[:-1], FileFormat=56)
#
#         wb.Close()
#
#         excel.Application.Quit()
#     # 释放资源
#     pythoncom.CoUninitialize()



# 后台自动打开和保存excel文档使文档公式计算生成两套值，方便下面之取值而不取公式
# def just_open(filename):
#     pythoncom.CoInitialize()
#     xlApp = Dispatch("Excel.Application")
#     xlApp.Visible = False
#     xlBook = xlApp.Workbooks.Open(filename)
#     xlBook.Save()
#     xlBook.Close()
#     pythoncom.CoUninitialize()


# if __name__ == '__main__':
#
#     getFiro2(fro,i,j)

