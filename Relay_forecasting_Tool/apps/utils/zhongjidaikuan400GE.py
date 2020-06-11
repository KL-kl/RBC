import xlrd
from xlutils.copy import copy
import os
import math


def band_400GE(node_dict,path,path2,BB1list,BB2list,BClist,v1,v2,v3,v4,v5,v6,v7,bandwidth1,v8,bandwidth2,v9,bandwidth3):
    book = xlrd.open_workbook(path, formatting_info=True)  # 打开文件，并且保留原格式
    sheet_list = book.sheet_names()
    book2 = copy(book)

    sheet = book.sheet_by_name('中继转发流量')  # 打开sheet页

    nodes = sheet.row_values(0)[1:]


    if '中继带宽400GE' in sheet_list:
        sheet_index = sheet_list.index('中继带宽400GE')
        sh = book2.get_sheet(sheet_index)
        for i in range(1, len(nodes) + 1):
            for j in range(1, len(nodes) + 1):
                flow = sheet.cell_value(i, j)
                if flow != '':

                    if (nodes[i - 1] in BB1list) and (nodes[j - 1] in BB1list):
                        if (nodes[i - 1][:2] == nodes[j - 1][:2]):
                            T1 = str(v7)+'*'+str(bandwidth1)+'GE'

                            sh.write(i, j, T1)
                        else:
                            T1 = math.ceil(flow / v1 / (400 * 1000))

                            sh.write(i, j, T1)


                    elif ((nodes[i - 1] in BB1list) and (nodes[j - 1] in BB2list)) or (
                            (nodes[i - 1] in BB2list) and (nodes[j - 1] in BB1list)):

                        T1 = math.ceil(flow / v2 / (400 * 1000))

                        sh.write(i, j, T1)

                    elif ((nodes[i - 1] in BClist) and (nodes[j - 1] in BB1list)) or (
                            (nodes[i - 1] in BB1list) and (nodes[j - 1] in BClist)):

                        T1 = math.ceil(flow / v3 / (400 * 1000))

                        sh.write(i, j, T1)

                    elif (nodes[i - 1] in BB2list) and (nodes[j - 1] in BB2list):
                        if (nodes[i - 1][:2] == nodes[j - 1][:2]):
                            T1 = str(v8)+'*'+str(bandwidth2) + 'GE'

                            sh.write(i, j, T1)
                        else:
                            T1 = math.ceil(flow / v4 / (400 * 1000))

                            sh.write(i, j, T1)

                    elif ((nodes[i - 1] in BClist) and (nodes[j - 1] in BB2list)) or (
                            (nodes[i - 1] in BB2list) and (nodes[j - 1] in BClist)):

                        T1 = math.ceil(flow / v5 / (400 * 1000))

                        sh.write(i, j, T1)

                    elif (nodes[i - 1] in BClist) and (nodes[j - 1] in BClist):
                        if (nodes[i - 1][:2] == nodes[j - 1][:2]):
                            T1 = str(v9)+'*'+str(bandwidth3) + 'GE'

                            sh.write(i, j, T1)
                        else:
                            T1 = math.ceil(flow / v6 / (400 * 1000))

                            sh.write(i, j, T1)



    else:
        worksheet = book2.add_sheet('中继带宽400GE', cell_overwrite_ok=True)

        for i in range(1, len(nodes) + 1):
            rows = sheet.cell(i, 0).value
            cols = sheet.cell(0, i).value
            worksheet.write(i, 0, rows)
            worksheet.write(0, i, cols)
            for j in range(1, len(nodes) + 1):
                flow = sheet.cell_value(i, j)
                if flow != '':

                    if (nodes[i - 1] in BB1list) and (nodes[j - 1] in BB1list):
                        if (nodes[i - 1][:2] == nodes[j - 1][:2]):
                            T1 = str(v7)+'*'+str(bandwidth1)+'GE'

                            worksheet.write(i, j, T1)
                        else:
                            T1 = math.ceil(flow / v1 / (400 * 1000))

                            worksheet.write(i, j, T1)


                    elif ((nodes[i - 1] in BB1list) and (nodes[j - 1] in BB2list)) or (
                            (nodes[i - 1] in BB2list) and (nodes[j - 1] in BB1list)):

                        T1 = math.ceil(flow / v2 / (400 * 1000))

                        worksheet.write(i, j, T1)

                    elif ((nodes[i - 1] in BClist) and (nodes[j - 1] in BB1list)) or (
                            (nodes[i - 1] in BB1list) and (nodes[j - 1] in BClist)):

                        T1 = math.ceil(flow / v3 / (400 * 1000))

                        worksheet.write(i, j, T1)

                    elif (nodes[i - 1] in BB2list) and (nodes[j - 1] in BB2list):
                        if (nodes[i - 1][:2] == nodes[j - 1][:2]):
                            T1 = str(v8)+'*'+str(bandwidth2) + 'GE'

                            worksheet.write(i, j, T1)
                        else:
                            T1 = math.ceil(flow / v4 / (400 * 1000))

                            worksheet.write(i, j, T1)

                    elif ((nodes[i - 1] in BClist) and (nodes[j - 1] in BB2list)) or (
                            (nodes[i - 1] in BB2list) and (nodes[j - 1] in BClist)):

                        T1 = math.ceil(flow / v5 / (400 * 1000))

                        worksheet.write(i, j, T1)

                    elif (nodes[i - 1] in BClist) and (nodes[j - 1] in BClist):
                        if (nodes[i - 1][:2] == nodes[j - 1][:2]):
                            T1 = str(v9)+'*'+str(bandwidth3) + 'GE'

                            worksheet.write(i, j, T1)
                        else:
                            T1 = math.ceil(flow / v6 / (400 * 1000))

                            worksheet.write(i, j, T1)


    if os.path.exists(path):
        os.remove(path)
        book2.save(path2)  # 保存新的excel，保存excel必须使用后缀名是.xls的，不是能是.xlsx的

        os.rename(path2, path)
    else:
        book2.save(path)
