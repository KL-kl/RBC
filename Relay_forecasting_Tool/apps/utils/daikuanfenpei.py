from django.http import JsonResponse
import os
import xlrd
from xlutils.copy import copy
from decimal import *


def bandwidth_max(path,path2):
    book = xlrd.open_workbook(path, formatting_info=True)  # 打开文件，并且保留原格式
    sheet = book.sheet_by_name('带宽')
    book2 = copy(book)
    sheets = book.sheet_names()
    nodes = sheet.row_values(0)[1:]

    remove_list = ['源节点', '目的节点', '目的   源 ', 'Metric', '合计']

    for i in remove_list:
        if i in nodes:
            nodes.remove(i)
        else:
            continue

    num = len(nodes)

    if '中继转发流量' in sheets:
        sh_index = sheets.index('中继转发流量')
        sh = book2.get_sheet(sh_index)

        for i in range(1, num + 1):
            for j in range(1, num + 1):
                sh.write(i,0,sheet.cell(i,0).value)
                sh.write(0, j, sheet.cell(0, j).value)
                s = sheet.cell(i, j).value
                e = sheet.cell(j, i).value

                if e > s:
                    bandweight = e
                else:
                    bandweight = s

                if i >= j:
                    if bandweight != 0 and bandweight != '':

                        sh.write(i, j, bandweight)

                    else:

                        sh.write(i, j, '')
    else:
        sh = book2.add_sheet('中继转发流量',cell_overwrite_ok=True)
        for i in range(1, num + 1):
            for j in range(1, num + 1):
                sh.write(i, 0, sheet.cell(i, 0).value)
                sh.write(0, j, sheet.cell(0, j).value)
                s = sheet.cell(i, j).value
                e = sheet.cell(j, i).value

                if e > s:
                    bandweight = e
                else:
                    bandweight = s

                if i >= j:
                    if bandweight != 0 and bandweight != '':

                        sh.write(i, j, bandweight)

                    else:

                        sh.write(i, j, '')


    if os.path.exists(path):
        os.remove(path)
        book2.save(path2)
        os.rename(path2, path)
    else:
        book2.save(path)  # 保存新的excel，保存excel必须使用后缀名是.xls的，不是能是.xlsx的


def bandwidth_allocation(pro_type,path1,path2):
    if pro_type == '2':

        path = os.path.abspath('excel/cmnetwin')
        book = xlrd.open_workbook(path1, formatting_info=True)

        sheet = book.sheet_by_name('点点路由')  # 打开sheet页
        sheet3 = book.sheet_by_name('直联')  # 打开sheet页
        sheet_list = book.sheet_names()
        if not os.path.join(path, 'CMNET网工程计算.xls'):
            return JsonResponse({'status':'fail','msg':'请先进行点到点流量计算！！'})

        b2 = xlrd.open_workbook(os.path.join(path, 'CMNET网工程计算.xls'))

        sheet1 = b2.sheet_by_name('点到点流量')

        nodes = sheet3.row_values(0)[1:]
        remove_list = ['源节点', '目的节点', '目的   源 ', 'Metric', '合计']

        for i in remove_list:
            if i in nodes:
                nodes.remove(i)
            else:
                continue
        num = len(nodes)
        book2 = copy(book)

        dic = {sheet.cell_value(i, 0): 0 for i in range(1, num * num + 1)}

        for i in range(1, num * num + 1):

            ll = sheet.cell_value(i, 0).split('-')
            st = ll[0]
            en = ll[1]
            if sheet.cell_value(i, 1) != 'NULL' and sheet.cell_value(i, 1) != '':  # 只有确实有路由才进行下面的计算
                # print(sheet.cell_value(i, 1))

                sameweight_num = int(sheet.cell(i, 1).value)  # 同权路由数
                # 各条同权路由负荷分担
                band = sheet1.cell_value(nodes.index(en) + 1, nodes.index(st) + 1)
                # print(type(band),band)
                if band != '':
                    sameweight_flow = Decimal(str(band)) / Decimal(str(sameweight_num))  # 同权路由
                    for same in range(1, sameweight_num + 1):
                        route = sheet.cell_value(i, 3 + same)

                        node_list = route.split('-')
                        str_size = len(node_list)

                        for str_l in range(1, str_size):
                            startnode = node_list[str_l - 1]
                            endnode = node_list[str_l]

                            s_e = startnode + '-' + endnode
                            for key in dic.keys():
                                if s_e == key:
                                    dic[key] = Decimal(str(dic[key])) + Decimal(str(sameweight_flow))

        if '带宽' in sheet_list:
            sheet_index = sheet_list.index('带宽')
            sh = book2.get_sheet(sheet_index)

            for k, v in dic.items():
                node = k.split('-')
                sh.write(nodes.index(node[1]) + 1, nodes.index(node[0]) + 1, v)

        else:

            worksheet = book2.add_sheet('带宽', cell_overwrite_ok=True)
            for i in range(1,num + 1):
                rows = sheet3.cell(i, 0).value
                cols = sheet3.cell(0, i).value
                worksheet.write(i, 0, rows)
                worksheet.write(0, i, cols)
                for k, v in dic.items():
                    node = k.split('-')
                    worksheet.write(nodes.index(node[1]) + 1, nodes.index(node[0]) + 1, v)

        if os.path.exists(path1):
            os.remove(path1)
            book2.save(path2)  # 保存新的excel，保存excel必须使用后缀名是.xls的，不是能是.xlsx的
            os.rename(path2, path1)
        else:
            book2.save(path1)
        bandwidth_max(path1,path2)
    elif pro_type == '1':
        path = os.path.abspath('excel/ipwin')
        book = xlrd.open_workbook(path1, formatting_info=True)

        sheet = book.sheet_by_name('点点路由')  # 打开sheet页
        sheet3 = book.sheet_by_name('直联')  # 打开sheet页
        sheet_list = book.sheet_names()

        flow = os.path.join(path,'IP网_点到点流量.xls')
        if not os.path.exists(flow):
            return JsonResponse({'status':'fail','msg':'请先导入点到点流量或进行点到点流量计算！！'})

        b2 = xlrd.open_workbook(os.path.join(path,'IP网_点到点流量.xls'))
        sheet1 = b2.sheet_by_index(0)

        nodes = sheet3.row_values(0)[1:]
        remove_list = ['源节点','目的节点','目的   源 ','Metric','合计']
        for i in remove_list:
            if i in nodes:
                nodes.remove(i)
            else:
                continue
        num = len(nodes)
        book2 = copy(book)

        dic = {sheet.cell_value(i, 0): 0 for i in range(1, num*num+1)}

        for i in range(1, num*num+1):

            ll = sheet.cell_value(i, 0).split('-')
            st = ll[0]
            en = ll[1]
            if sheet.cell_value(i, 1) != 'NULL' and sheet.cell_value(i, 1) != '':  # 只有确实有路由才进行下面的计算
                # print(sheet.cell_value(i, 1))
                sameweight_num = int(sheet.cell_value(i, 1))  # 同权路由数
                # 各条同权路由负荷分担
                band = sheet1.cell_value(nodes.index(en) + 1, nodes.index(st) + 1)
                # print(type(band),band)
                if band != '':
                    sameweight_flow = Decimal(str(band)) / Decimal(str(sameweight_num))  # 同权路由
                    for same in range(1, sameweight_num + 1):
                        route = sheet.cell_value(i, 3 + same)

                        node_list = route.split('-')
                        str_size = len(node_list)

                        for str_l in range(1, str_size):
                            startnode = node_list[str_l - 1]
                            endnode = node_list[str_l]

                            s_e = startnode+'-'+endnode
                            for key in dic.keys():
                                if s_e == key:
                                    dic[key] = Decimal(str(dic[key])) + Decimal(str(sameweight_flow))

        if '带宽' in sheet_list:
            sheet_index = sheet_list.index('带宽')
            sh = book2.get_sheet(sheet_index)

            for k, v in dic.items():
                node = k.split('-')
                sh.write(nodes.index(node[1])+1, nodes.index(node[0])+1,v)

        else:
            worksheet = book2.add_sheet('带宽',cell_overwrite_ok=True)
            for i in range(1,num + 1):
                rows = sheet3.cell(i, 0).value
                cols = sheet3.cell(0, i).value
                worksheet.write(i, 0, rows)
                worksheet.write(0, i, cols)
                for k, v in dic.items():
                    node = k.split('-')
                    worksheet.write(nodes.index(node[1])+1, nodes.index(node[0])+1, v)

        if os.path.exists(path1):
            os.remove(path1)
            book2.save(path2)  # 保存新的excel，保存excel必须使用后缀名是.xls的，不是能是.xlsx的
            os.rename(path2, path1)
        else:
            book2.save(path1)
        bandwidth_max(path1,path2)
