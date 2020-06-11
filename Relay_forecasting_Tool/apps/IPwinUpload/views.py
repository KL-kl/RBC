import os
from django.shortcuts import render, HttpResponse
from django.utils.encoding import escape_uri_path
from django.db import connection
from django.http import JsonResponse
import re
import xlrd
import xlwt
from xlutils.copy import copy
import logging
import json
import math
import numpy as np
from Relay_forecasting_Tool import settings
from IPwinUpload.models import IpNodes_detail, IP_TE
from CMNETwinUpload.models import ProjectInfo
from utils.hebin import get_exce, get_name, open_exce
from utils.fault_polling import *
from utils.fault_polling_min import *
from utils.sort_path import *
from utils.daikuanfenpei import bandwidth_allocation
from utils.excel_style import *
from utils.calculateTable import getMaxFiro,getFiro2

# Create your views here.

# 自定义表格样式
logger = logging.getLogger('log')

# 自定义表格样式
style = xlwt.easyxf('align: wrap on, vert centre, horiz center;'
                    'borders:left 1,right 1,top 1,bottom 1')

style2 = xlwt.easyxf('font: bold on,name Calibri;align: wrap on, vert centre, horiz center;'
                     'borders:left 1,right 1,top 1,bottom 1')

style3 = xlwt.easyxf('font: name Calibri;align: wrap on, vert centre, horiz center;'
                     'borders:left 1,right 1,top 1,bottom 1')

he_style3 = easyxf('pattern: pattern solid, fore_colour ice_blue; '
                   'font: bold on,name Calibri,height 260;'
                   'align: wrap on, vert centre, horiz center;'
                   'borders:left 2,right 2,top 2,bottom 2')
gz_style = xlwt.easyxf('pattern: pattern solid, fore_colour red;')

# TE style
te_style = xlwt.easyxf('pattern: pattern solid, fore_colour gold;')
te_style2 = xlwt.easyxf('pattern: pattern solid, fore_colour 44;')


def IPwin(request):
    return render(request, 'IPwin.html')


def IPwin_2(request):
    return render(request, 'ip_win.html')

#节点信息
def find_node():
    c = []
    b = []
    a = []
    data1 = IpNodes_detail.objects.filter(part2__in=['CR/BR', 'CR']).values('province2', 'part2', 'node2')

    data2 = IpNodes_detail.objects.filter(part2='CR/BR').values('province2', 'part2', 'node2')
    data3 = IpNodes_detail.objects.filter(part2='BR').values('province2', 'part2', 'node2')

    for i in range(len(data1)):
        c.append(data1[i]['province2'])
        data1[i].update(
            {'nick_pro': data1[i]['province2'][:1] + 'C' + str(c + 1) for c in
             range(c.count(data1[i]['province2']))})

    for i in range(len(data2)):
        b.append(data2[i]['province2'])

        data2[i].update(
            {'nick_pro': data2[i]['province2'][:2] + str(b + 1) for b in range(b.count(data2[i]['province2']))})

    for i in range(len(data3)):
        a.append(data3[i]['province2'])

        if a.count(data3[i]['province2']) > 2:

            data3[i].update(
                {'nick_pro': data3[i]['node2'][:2] + str(a - 1) for a in range(a.count(data3[i]['province2']))})
        else:
            data3[i].update(
                {'nick_pro': data3[i]['province2'][:2] + str(a + 1) for a in range(a.count(data3[i]['province2']))})
    d1 = {}
    d2 = {}
    d3 = {}
    for i in data1:
        # s = '北京CR/BR1'
        # num = re.sub("\D", "", i['node2'])  # 1
        # res = i['node2'].split('/')  # ['北京CR', 'BR1']
        # r1 = res[0] + num    # 北京CR1
        if len(i['node2']) < 6:
            d1[i['nick_pro']] = i['node2']
        else:
            d1[i['nick_pro']] = (i['node2'].split('/'))[0] + re.sub("\D", "", i['node2'])
    for i in data2:
        # s = '北京CR/BR1'
        # str = re.sub("[A-Za-z0-9\/]", "", i['node2']) #北京
        # res = i['node2'].split('/') # ['北京CR', 'BR1']
        # r2 = str + res[1] # 北京BR1

        d2[i['nick_pro']] = re.sub("[A-Za-z0-9\/]", "", i['node2']) + (i['node2'].split('/'))[1]

    for i in data3:
        d3[i['nick_pro']] = i['node2']

    d1.update(d2)
    d1.update(d3)

    return d1  # 链路字典


# 上传IP网节点数据入库
def upload_Ifile(request):
    if request.method == 'POST':

        pro_name = request.POST.get('pro_name1')
        myFile = request.FILES.get('node_file1')
        if not all([pro_name, myFile]):
            # 有数据为空
            return JsonResponse({'status': 'err', 'msg': '参数不能为空!'})

        path = settings.MEDIA_ROOT + 'ipwin/nodes/'  # 上传文件的保存路径，可以自己指定任意的路径

        if not os.path.exists(path):
            os.makedirs(path)

        # 创建模板
        model_path = os.path.abspath('models')
        if not os.path.exists(model_path):
            os.makedirs(model_path)

        with open(os.path.join(path, myFile.name), 'wb+') as f:
            for chunk in myFile.chunks():
                f.write(chunk)

        node_list = IpNodes_detail.objects.all()
        if len(node_list) > 0:
            cursor = connection.cursor()  # 建立游标

            cursor.execute("TRUNCATE TABLE e_ipnodes_detail")
            cursor.execute("ALTER TABLE e_ipnodes_detail AUTO_INCREMENT = 1")

            cursor.close()  # 关闭游标

            excel = xlrd.open_workbook(path + myFile.name)
            sheet = excel.sheet_by_index(1)
            nrows = sheet.nrows
        else:
            cursor = connection.cursor()  # 建立游标

            cursor.execute("ALTER TABLE e_ipnodes_detail AUTO_INCREMENT = 1")

            cursor.close()  # 关闭游标
            excel = xlrd.open_workbook(path + myFile.name)
            sheet = excel.sheet_by_index(1)
            nrows = sheet.nrows

        for i in range(2, nrows):
            inode = IpNodes_detail()
            inode.province1 = sheet.cell(i, 1).value
            inode.node1 = sheet.cell(i, 2).value
            inode.office_address1 = sheet.cell(i, 3).value
            inode.building_no1 = sheet.cell(i, 4).value
            inode.floor1 = sheet.cell(i, 5).value
            inode.room_num1 = sheet.cell(i, 6).value
            inode.plane1 = sheet.cell(i, 7).value
            inode.part1 = sheet.cell(i, 8).value
            inode.network_level1 = sheet.cell(i, 9).value
            inode.province2 = sheet.cell(i, 10).value
            inode.node2 = sheet.cell(i, 11).value
            inode.office_address2 = sheet.cell(i, 12).value
            inode.building_no2 = sheet.cell(i, 13).value
            inode.floor2 = sheet.cell(i, 14).value
            inode.room_num2 = sheet.cell(i, 15).value
            inode.plane2 = sheet.cell(i, 16).value
            inode.part2 = sheet.cell(i, 17).value
            inode.network_level2 = sheet.cell(i, 18).value
            inode.program = sheet.cell(i, 19).value

            inode.save()

        pro = ProjectInfo()
        pro.id = 1
        pro.name = pro_name
        pro.pro_type = 1
        pro.nodefile = os.path.join(path, myFile.name)
        pro.createBy_id = request.user.id
        pro.save()

        node_dict = find_node()
        ll = list(node_dict.keys())

        work_book1 = xlwt.Workbook(encoding='utf-8')
        sheet1 = work_book1.add_sheet('Metric')
        sheet1.write(0, 0, 'Metric')

        work_book2 = xlwt.Workbook(encoding='utf-8')
        sheet2 = work_book2.add_sheet('传输距离')
        sheet2.write(0, 0, '传输距离')

        for i in range(len(ll)):
            sheet1.write(0, i + 1, ll[i])
            sheet1.write(i + 1, 0, ll[i])
            sheet2.write(0, i + 1, ll[i])
            sheet2.write(i + 1, 0, ll[i])

        model_path = os.path.abspath('models')
        if not os.path.exists(model_path):
            os.mkdir(model_path)
        work_book1.save(os.path.join(model_path, 'IP网_metric.xls'))
        work_book2.save(os.path.join(model_path, 'IP网_传输距离.xls'))

        return JsonResponse({'status': 'ok', 'msg': '上传成功！！！'})

#点到点流量
def ptop_flow():
    path = os.path.abspath('excel/ipwin')
    flow_path = os.path.join(path, 'combine.xls')
    if os.path.exists(flow_path):
        wb = xlrd.open_workbook(flow_path)
        flow_sh = wb.sheet_by_name('所有业务需求汇总')
        nodes = flow_sh.row_values(0)[1:]
        remove_list = ['源节点', '目的节点', '合计', '节点发送带宽和']
        for i in remove_list:
            if i in nodes:
                nodes.remove(i)
        num = len(nodes)
        flow_dict = {}
        for i in range(1, num + 1):
            for j in range(1, num + 1):
                flow = flow_sh.cell(i, j).value
                if flow == 0 or flow == '':
                    continue
                else:
                    s = nodes[i - 1][3:]
                    e = nodes[j - 1][3:]
                    flow_dict[s + '-' + e] = flow
        # print(flow_dict)

        d1 = find_node()
        # print(d1)

        wb = xlrd.open_workbook(os.path.join(path, 'IP网_直联.xls'))
        ws = wb.sheet_by_index(0)
        nick_node = ws.row_values(0)[1:]
        num = len(nick_node)

        exc = xlwt.Workbook()
        sheet = exc.add_sheet('点到点流量', cell_overwrite_ok=True)
        sheet.write(0, 0, 'Mb/s')

        for i in range(1, num + 1):
            sheet.write(i, 0, d1[nick_node[i - 1]])
            sheet.write(0, i, d1[nick_node[i - 1]])
            for j in range(1, num + 1):
                for k, v in flow_dict.items():

                    node = k.split('-')

                    st = node[0] + 'BR'
                    en = node[1] + 'BR'
                    if (st in d1[nick_node[i - 1]]) and (en in d1[nick_node[j - 1]]):
                        sheet.write(i, j, v / 4)
                        sheet.write(j, i, v / 4)

        if os.path.exists(os.path.join(path, 'IP网_点到点流量.xls')):

            try:
                os.remove(os.path.join(path, 'IP网_点到点流量.xls'))
                exc.save(os.path.join(path, 'IP网_点到点流量.xls'))
            except Exception as e:
                logger.error(e)
                return JsonResponse({'status': 'err', 'msg': '文件被打开，计算失败，请先关闭文件重新计算！！！！'})
        else:
            exc.save(os.path.join(path, 'IP网_点到点流量.xls'))


# 导入多张业务表计算点到点流量
def up_maxfile(request):
    if request.method == 'POST':

        xs = request.POST.get('xs')
        pro_type = request.POST.get('pro_type')
        path = settings.MEDIA_ROOT + 'ipwin/max/'  # 上传文件的保存路径，可以自己指定任意的路径
        if not os.path.exists(path):
            os.makedirs(path)

        path2 = os.path.abspath('excel/ipwin')
        if not os.path.exists(path2):
            os.makedirs(path2)
        if not os.path.exists(path):
            os.makedirs(path)

        files = request.FILES.getlist('myfiles')
        file_li = []
        for f in files:
            file_li.append(f.name)
            destination = open(path + f.name, 'wb+')
            for chunk in f.chunks():
                destination.write(chunk)
            destination.close()

        pro = ProjectInfo.objects.get(id=1)
        pro.xs = xs
        # pro.businessfile = files
        pro.save()

        all_exce = get_exce(pro_type)
        # 得到要合并的所有exce表格数据
        if (all_exce == 0):
            os.system('pause')
            exit()
            return JsonResponse({'status': 'err', 'msg': '该目录下无.xlsx文件！请检查您输入的目录是否有误！！！！'})

        name = get_name(pro_type)
        wb_result = xlwt.Workbook()  # 新建一个文件，用来保存结果
        wb_result2 = xlwt.Workbook()
        sheet_result2 = wb_result2.add_sheet('各省流量汇总', cell_overwrite_ok=True)
        sheet_result2.write(0, 0, '序号', style2)
        sheet_result2.write(0, 1, '省份', style2)
        sheet_result2.write(0, 2 + len(all_exce), '合计', style2)
        for i in range(len(all_exce)):
            sheet_result2.write(0, 2 + i, name[i][:-4], style2)

        for exce in all_exce:
            sheet_name = name[all_exce.index(exce)][:-4]
            fh = open_exce(exce)

            sheet_result = wb_result.add_sheet(sheet_name, cell_overwrite_ok=True)
            sheet_pri = fh.sheet_by_index(0)  # 通过index获取每个sheet
            row_0 = sheet_pri.row_values(0)  # 获取第一行的值
            col_0 = sheet_pri.col_values(0)  # 获取第一列的值
            for i, key in enumerate(row_0):  # 写入新Excel表的第一行
                sheet_result.write(0, i, key)
            for i, key in enumerate(col_0):  # 写入新Excel表的第一列
                sheet_result.write(i, 0, key)

            for i, key in enumerate(row_0[1:]):  # 写入新Excel表的第一行
                ll = key.split('-')

                if len(ll) < 2:
                    sheet_result2.write(i + 1, 0, '', style)
                    sheet_result2.write(i + 1, 1, ll[0], style3)
                else:
                    st = ll[0]
                    en = ll[1]
                    # print(st, en)  # 02 天津
                    sheet_result2.write(i + 1, 0, int(st), style)
                    sheet_result2.write(i + 1, 1, en, style)

            ncols = sheet_pri.ncols
            nrows = sheet_pri.nrows
            for i in range(1, nrows):
                for j in range(1, ncols):
                    ban = sheet_pri.cell_value(i, j)
                    sheet_result.write(i, j, ban)
            wb_result.save(os.path.join(path2, 'combine.xls'))

            he = 0

            for i in range(1, ncols - 1):
                ban1 = sheet_pri.row_values(i)[1:i]
                ban2 = sheet_pri.col_values(i)[i:-1]

                total = sum(ban1 + ban2)

                he += total
                sheet_result2.write(i, all_exce.index(exce) + 2, total, style)

            sheet_result2.write(ncols - 1, all_exce.index(exce) + 2, he, style)

        p1 = os.path.join(path2, 'IP承载网各省业务流量需求汇总.xls')
        p2 = os.path.join(path2, 'IP承载网各省业务流量需求汇总2.xls')
        if os.path.exists(p1):
            os.remove(p1)
            wb_result2.save(p2)
            os.rename(p2, p1)
        else:
            wb_result2.save(p1)

        b = xlrd.open_workbook(p1, formatting_info=True)
        b2 = copy(b)
        sheet_pri = b.sheet_by_index(0)  # 通过index获取每个sheet
        nrows = sheet_pri.nrows
        sh = b2.get_sheet(0)

        for i in range(1, nrows):
            row = sheet_pri.row_values(i)[2:-1]  # 获取第一行的值
            if '' in row:
                continue
            sh.write(i, len(all_exce) + 2, sum(row), style)

        if os.path.exists(p1):
            os.remove(p1)
            b2.save(p2)
            os.rename(p2, p1)
        else:
            b2.save(p1)

        if os.path.exists(os.path.join(path2, 'combine.xls')):

            wb = open_exce(os.path.join(path2, 'combine.xls'))
            sh = wb.sheet_by_index(0)
            ncols = sh.ncols  # Excel列的数目  原Excel和目标Excel的列表的长度相同
            nrows = sh.nrows
            test = [[0 for c in range(ncols - 1)] for r in range(nrows - 1)]

            for sheetName in wb.sheet_names():
                sheet = wb.sheet_by_name(sheetName)

                for i in range(1, nrows):
                    for j in range(1, ncols):

                        ban = sheet.cell(i, j).value

                        if ban == '':
                            ban = 0

                        test[i - 1][j - 1] = test[i - 1][j - 1] + ban
            # print(test)

            book2 = copy(wb)
            ws = book2.add_sheet('所有业务需求汇总', cell_overwrite_ok=True)
            row_0 = sh.row_values(0)  # 获取第一行的值
            col_0 = sh.col_values(0)  # 获取第一列的值
            for i, key in enumerate(row_0):  # 写入新Excel表的第一行
                ws.write(0, i, key)
            for i, key in enumerate(col_0):  # 写入新Excel表的第一列
                ws.write(i, 0, key)
            x = np.array(test)
            # 输出数组的行和列数
            for i in range(x.shape[0]):
                for j in range(x.shape[1]):
                    ws.write(i + 1, j + 1, x[i][j])

            p = os.path.join(path2, 'combine.xls')
            p2 = os.path.join(path2, 'new_combine.xls')
            if os.path.exists(p):
                os.remove(p)
                book2.save(p2)
                os.rename(p2, p)
            else:
                book2.save(p)

        ptop_flow()

        return JsonResponse({'status': 'ok', 'msg': '上传成功！！！'})


def up_metric1(request):
    if request.method == 'POST':
        metric_file = request.FILES.get('metric_file')
        if not metric_file:
            return JsonResponse({'status': 'err', 'msg': '文件不存在！！！'})
        else:
            path = settings.MEDIA_ROOT + 'ipwin/metric/'  # 上传文件的保存路径，可以自己指定任意的路径
            path2 = os.path.abspath('excel/ipwin')
            if not os.path.exists(path):
                os.makedirs(path)
            if not os.path.exists(path2):
                os.makedirs(path2)
            with open(os.path.join(path, metric_file.name), 'wb+') as f:
                for chunk in metric_file.chunks():
                    f.write(chunk)

            excel = xlrd.open_workbook(path + metric_file.name)
            sheet = excel.sheet_by_index(0)
            nodes = sheet.row_values(0)[1:]

            wb = xlwt.Workbook(encoding='utf-8')
            sh = wb.add_sheet('直联', cell_overwrite_ok=True)
            sh.write(0, 0, 'Metric')

            nrows = sheet.nrows
            for i in range(1, nrows):
                sh.write(i, 0, nodes[i - 1])
                sh.write(0, i, nodes[i - 1])
                for j in range(1, nrows):

                    m = sheet.cell(i, j).value
                    if m == '' or m == 0:
                        m = ''

                    sh.write(i, j, m)
                    sh.write(j, i, m)

            wb.save(os.path.join(path2, 'IP网_直联.xls'))

        return JsonResponse({'status': 'ok', 'msg': '上传成功！！！'})


def up_metric2(request):
    if request.method == 'POST':
        myFile = request.FILES.get('distance_file')
        js1 = request.POST.get('js1')
        zs1 = request.POST.get('zs1')
        js2 = request.POST.get('js2')
        zs2 = request.POST.get('zs2')
        js3 = request.POST.get('js3')
        zs3 = request.POST.get('zs3')

        if not all([myFile, js1, zs1, js2, zs2, js3, zs3]):
            # 有数据为空
            return JsonResponse({'status': 'err', 'msg': '参数不能为空!'})

        path = settings.MEDIA_ROOT + 'ipwin/distance/'  # 上传文件的保存路径，可以自己指定任意的路径
        if not os.path.exists(path):
            os.makedirs(path)

        with open(os.path.join(path, myFile.name), 'wb+') as f:
            for chunk in myFile.chunks():
                f.write(chunk)

        excel = xlrd.open_workbook(path + myFile.name)
        sheet = excel.sheet_by_index(0)

        nrows = sheet.nrows

        node_list = IpNodes_detail.objects.all().values()
        if len(node_list) > 0:
            dic = find_node()
            nodes = list(dic.keys())

            cr = []
            br = []
            for n in nodes:
                if "C" in n:
                    cr.append(n)
                else:
                    br.append(n)

            work_book1 = xlwt.Workbook(encoding='utf-8')
            sheet1 = work_book1.add_sheet('Metric', cell_overwrite_ok=True)
            sheet1.write(0, 0, 'Metric')

            work_book2 = xlwt.Workbook(encoding='utf-8')
            sheet2 = work_book2.add_sheet('直联', cell_overwrite_ok=True)

            for i in range(1, nrows):
                sheet1.write(i, 0, nodes[i - 1])
                sheet1.write(0, i, nodes[i - 1])
                for j in range(1, nrows):

                    m = sheet.cell(i, j).value
                    if m == '' or m == 0:
                        m = 0
                    else:

                        # if (nodes[i - 1][:2] == nodes[j - 1][:2]) and (nodes[i-1] in br) and (nodes[j-1] in br):
                        #     sheet1.write(i, j, eval(js3) - eval(zs3))
                        #     zs3 += 1
                        # elif (nodes[i - 1][:2] == nodes[j - 1][:2]) and (
                        #         (nodes[i - 1] in br) and (nodes[j - 1] in br)):
                        #     m = eval(js1) + m / eval(zs1)
                        #     sheet1.write(i, j, m)
                        # elif ((nodes[i - 1] in cr) and (nodes[j - 1] in br)) or (
                        #         (nodes[i - 1] in br) and (nodes[j - 1] in cr)):
                        #     m = eval(js2) + m / eval(zs2)
                        #     sheet1.write(i, j, m)
                        if (nodes[i - 1] in cr) and (nodes[j - 1] in cr):
                            # metric
                            sheet1.write(i, j, eval(js1) + m / eval(zs1))
                            # 生成计算需要的直联表
                            sheet2.write(i, j, eval(js1) + m / eval(zs1))
                            sheet2.write(j, i, eval(js1) + m / eval(zs1))


                        elif ((nodes[i - 1] in cr) and (nodes[j - 1] in br)) or (
                                (nodes[i - 1] in br) and (nodes[j - 1] in cr)):

                            sheet1.write(i, j, eval(js2) + m / eval(zs2))

                            sheet2.write(i, j, eval(js2) + m / eval(zs2))
                            sheet2.write(j, i, eval(js2) + m / eval(zs2))

                        elif (nodes[i - 1][:2] == nodes[j - 1][:2]) and (
                                (nodes[i - 1] in br) and (nodes[j - 1] in br)):

                            sheet1.write(i, j, eval(js3) - eval(zs3))

                            sheet2.write(i, j, eval(js3) - eval(zs3))
                            sheet2.write(j, i, eval(js3) - eval(zs3))

            metric_path = os.path.abspath('excel/ipwin/metric')
            if not os.path.exists(metric_path):
                os.makedirs(metric_path)

            if os.path.exists(os.path.join(metric_path, 'IP网_metric.xls')):
                work_book1.save(
                    os.path.join(metric_path, 'IP网_metric2.xls'))  # 保存新的excel，保存excel必须使用后缀名是.xls的，不是能是.xlsx的
                os.remove(os.path.join(metric_path, 'IP网_metric.xls'))
                os.rename(os.path.join(metric_path, 'IP网_metric2.xls'), os.path.join(metric_path, 'IP网_metric.xls'))
            else:
                work_book1.save(os.path.join(metric_path, 'IP网_metric.xls'))

            # 保存计算直联
            path1 = os.path.abspath('excel/ipwin')
            if not os.path.exists(path1):
                os.makedirs(path1)

            if os.path.exists(os.path.join(path1, 'IP网_直联.xls')):
                work_book2.save(
                    os.path.join(path1, 'IP网_直联2.xls'))  # 保存新的excel，保存excel必须使用后缀名是.xls的，不是能是.xlsx的
                os.remove(os.path.join(path1, 'IP网_直联.xls'))
                os.rename(os.path.join(path1, 'IP网_直联2.xls'), os.path.join(path1, 'IP网_直联.xls'))
            else:
                work_book2.save(os.path.join(path1, 'IP网_metric.xls'))

            return JsonResponse({'status': 'ok', 'msg': '上传成功！'})


def up_metric3(request):
    if request.method == 'POST':
        myFile = request.FILES.get('distance_file2')
        zs1 = request.POST.get('zs1_1')
        zs2 = request.POST.get('zs2_2')
        js3 = request.POST.get('js3_3')
        zs3 = request.POST.get('zs3_3')

        if not myFile:
            return JsonResponse({'status': 'err', 'msg': '上传失败！'})
        path = settings.MEDIA_ROOT + 'ipwin/distance/'  # 上传文件的保存路径，可以自己指定任意的路径
        if not os.path.exists(path):
            os.makedirs(path)

        with open(os.path.join(path, myFile.name), 'wb+') as f:
            for chunk in myFile.chunks():
                f.write(chunk)

        excel = xlrd.open_workbook(os.path.join(path,myFile.name))
        sheet = excel.sheet_by_index(0)

        nrows = sheet.nrows

        node_list = IpNodes_detail.objects.all()
        if len(node_list) > 0:
            dic = find_node()
            nodes = list(dic.keys())

            cr = []
            br = []
            for n in nodes:
                if "C" in n:
                    cr.append(n)
                else:
                    br.append(n)

            work_book1 = xlwt.Workbook(encoding='utf-8')
            sheet1 = work_book1.add_sheet('Metric', cell_overwrite_ok=True)
            work_book2 = xlwt.Workbook(encoding='utf-8')
            sheet2 = work_book1.add_sheet('直联', cell_overwrite_ok=True)
            sheet1.write(0, 0, 'Metric')

            for i in range(1, nrows):
                sheet1.write(i, 0, nodes[i - 1])
                sheet1.write(0, i, nodes[i - 1])
                for j in range(1, nrows):

                    m = sheet.cell(i, j).value
                    if m == '' or m == 0:
                        m = 0
                    else:
                        if (nodes[i - 1] in cr) and (nodes[j - 1] in cr):
                            # metric
                            sheet1.write(i, j, m / eval(zs1))
                            # 生成计算需要的直联表
                            sheet2.write(i, j, m / eval(zs1))
                            sheet2.write(j, i, m / eval(zs1))


                        elif ((nodes[i - 1] in cr) and (nodes[j - 1] in br)) or (
                                (nodes[i - 1] in br) and (nodes[j - 1] in cr)):

                            sheet1.write(i, j, m / eval(zs2))

                            sheet2.write(i, j, m / eval(zs2))
                            sheet2.write(j, i, m / eval(zs2))

                        elif (nodes[i - 1][:2] == nodes[j - 1][:2]) and (
                                (nodes[i - 1] in br) and (nodes[j - 1] in br)):
                            sheet1.write(i, j, eval(js3) - eval(zs3))

                            sheet2.write(i, j, eval(js3) - eval(zs3))
                            sheet2.write(j, i, eval(js3) - eval(zs3))

            metric_path = os.path.abspath('excel/ipwin/metric')
            if not os.path.exists(metric_path):
                os.mkdir(metric_path)

            if os.path.exists(os.path.join(metric_path, 'IP网_metric.xls')):
                work_book1.save(
                    os.path.join(metric_path, 'IP网_metric2.xls'))  # 保存新的excel，保存excel必须使用后缀名是.xls的，不是能是.xlsx的
                os.remove(os.path.join(metric_path, 'IP网_metric.xls'))
                os.rename(os.path.join(metric_path, 'IP网_metric2.xls'), os.path.join(metric_path, 'IP网_metric.xls'))
            else:
                work_book1.save(os.path.join(metric_path, 'IP网_metric.xls'))

                # 保存计算直联
            path1 = os.path.abspath('excel/ipwin')
            if not os.path.exists(path1):
                os.makedirs(path1)

            if os.path.exists(os.path.join(path1, 'IP网_直联.xls')):
                work_book2.save(
                    os.path.join(path1, 'IP网_直联2.xls'))  # 保存新的excel，保存excel必须使用后缀名是.xls的，不是能是.xlsx的
                os.remove(os.path.join(path1, 'IP网_直联.xls'))
                os.rename(os.path.join(path1, 'IP网_直联2.xls'), os.path.join(path1, 'IP网_直联.xls'))
            else:
                work_book2.save(os.path.join(path1, 'IP网_metric.xls'))

                return JsonResponse({'status': 'ok', 'msg': '上传成功！'})

            return JsonResponse({'status': 'ok', 'msg': '上传成功！'})


def up_metric4(request):
    if request.method == 'POST':
        js1 = request.POST.get('js_1')
        js2 = request.POST.get('js_2')
        js3 = request.POST.get('js_3')
        zs3 = request.POST.get('zs_3')
        print(js1,js2,js3,zs3)
        if not all([js1,js2,js3,zs3]):
            return JsonResponse({'status': 'err', 'msg': '参数不能为空!'})

        # 保存计算metric
        metric_path = os.path.abspath('excel/ipwin/metric')
        if not os.path.exists(metric_path):
            os.mkdir(metric_path)

        node_list = IpNodes_detail.objects.all()
        if len(node_list) > 0:
            dic = find_node()
            # print(dic)

            work_book1 = xlwt.Workbook(encoding='utf-8')
            sheet1 = work_book1.add_sheet('Metric', cell_overwrite_ok=True)
            # work_book2 = xlwt.Workbook(encoding='utf-8')
            # sheet2 = work_book1.add_sheet('直联', cell_overwrite_ok=True)
            sheet1.write(0, 0, 'Metric')
            nodes = list(dic.keys())

            cr = []
            br = []

            for n in nodes:
                if 'C' in n:
                    cr.append(n)
                else:
                    br.append(n)

            for i in range(1, len(nodes)):
                sheet1.write(i, 0, nodes[i - 1])
                sheet1.write(0, i, nodes[i - 1])
                for j in range(1, len(nodes)):
                    if nodes[i - 1] == nodes[j - 1]:
                        continue
                    else:
                        if (nodes[i - 1] in cr) and (nodes[j - 1] in cr):
                            # metric
                            sheet1.write(i, j, eval(js1))

                            # 生成计算需要的直联表
                            # sheet2.write(i, j, m / eval(zs1))
                            # sheet2.write(j, i, m / eval(zs1))


                        elif ((nodes[i - 1] in cr) and (nodes[j - 1] in br)) or (
                                (nodes[i - 1] in br) and (nodes[j - 1] in cr)):

                            sheet1.write(i, j, eval(js2))


                            # sheet2.write(i, j, m / eval(zs2))
                            # sheet2.write(j, i, m / eval(zs2))

                        elif (nodes[i - 1][:2] == nodes[j - 1][:2]) and (
                                (nodes[i - 1] in br) and (nodes[j - 1] in br)):
                            sheet1.write(i, j, eval(js3) - eval(zs3))


                            # sheet2.write(i, j, eval(js3) - eval(zs3))
                            # sheet2.write(j, i, eval(js3) - eval(zs3))




            if os.path.exists(os.path.join(metric_path, 'IP网_metric.xls')):
                work_book1.save(
                    os.path.join(metric_path, 'IP网_metric2.xls'))  # 保存新的excel，保存excel必须使用后缀名是.xls的，不是能是.xlsx的
                os.remove(os.path.join(metric_path, 'IP网_metric.xls'))
                os.rename(os.path.join(metric_path, 'IP网_metric2.xls'), os.path.join(metric_path, 'IP网_metric.xls'))
            else:
                work_book1.save(os.path.join(metric_path, 'IP网_metric.xls'))

            # 保存计算直联
            # path1 = os.path.abspath('excel/ipwin')
            # if not os.path.exists(path1):
            #     os.makedirs(path1)
            #
            # if os.path.exists(os.path.join(path1, 'IP网_直联.xls')):
            #     work_book2.save(
            #         os.path.join(path1, 'IP网_直联2.xls'))  # 保存新的excel，保存excel必须使用后缀名是.xls的，不是能是.xlsx的
            #     os.remove(os.path.join(path1, 'IP网_直联.xls'))
            #     os.rename(os.path.join(path1, 'IP网_直联2.xls'), os.path.join(path1, 'IP网_直联.xls'))
            # else:
            #     work_book2.save(os.path.join(path1, 'IP网_直联.xls'))
            #
            return JsonResponse({'status': 'ok', 'msg': '计算成功！'})


def down_metric1(request):
    if request.method == 'GET':
        path = os.path.join(os.path.abspath('excel/ipwin/metric'),'/IP网_metric.xls')

        wb = xlrd.open_workbook(path, formatting_info=True)
        wb2 = copy(wb)

        response = HttpResponse(content_type='application/octet-stream')
        response['Content-Disposition'] = 'attachment; filename="{0}"'.format(escape_uri_path('IP网_metric.xls'))
        wb2.save(response)
        return response


# te 上传
def te_upload(request):
    if request.method == 'POST':
        te_file = request.FILES.get('te_file')

        if not te_file:
            # 有数据为空
            return JsonResponse({'status': 'err', 'msg': '参数不能为空!'})

        path = settings.MEDIA_ROOT + 'ipwin/te/'  # 上传文件的保存路径，可以自己指定任意的路径
        if not os.path.exists(path):
            os.makedirs(path)

        with open(os.path.join(path, te_file.name), 'wb+') as f:
            for chunk in te_file.chunks():
                f.write(chunk)

        te_list = IP_TE.objects.all()
        if len(te_list) > 0:
            IP_TE.objects.all().delete()

            excel = xlrd.open_workbook(path + te_file.name)
            sheet = excel.sheet_by_index(0)
            nrows = sheet.nrows
            sheet2 = excel.sheet_by_index(1)
            nrows2 = sheet2.nrows
        else:
            excel = xlrd.open_workbook(path + te_file.name)
            sheet = excel.sheet_by_index(0)
            nrows = sheet.nrows
            sheet2 = excel.sheet_by_index(1)
            nrows2 = sheet2.nrows

        for i in range(1, nrows):
            te = IP_TE()
            te.main_route = sheet.cell(i, 0).value
            te.standby_route = sheet.cell(i, 1).value
            te.save()

        for i in range(1, nrows2):
            te = IP_TE()
            te.main_route = sheet2.cell(i, 0).value
            te.standby_route = sheet2.cell(i, 1).value
            te.save()

        return JsonResponse({'status': 'ok', 'msg': '上传成功'})
    else:
        return JsonResponse({'status': 'fail', 'msg': '上传失败'})


# te_计算
def te_sortPath(request):
    if request.method == 'GET':

        d1 = find_node()

        te_CR = {}
        te_pro = {}
        te_data = IP_TE.objects.filter(main_route__icontains='CR').values()
        te_data1 = IP_TE.objects.all().exclude(main_route__icontains='CR').values()
        for i in te_data:
            te_CR[i['main_route']] = i['standby_route']

        for i in te_data1:
            te_pro[i['main_route']] = i['standby_route']

        # print(te_CR) #{'北京CR1-北京CR2': '北京CR1-上海CR1-上海CR2-北京CR2', '上海CR1-上海CR2': '上海CR1-北京CR1-北京CR2-上海CR2', '广州CR1-广州CR2': '广州CR1-武汉CR1-武汉CR2-广州CR2', '沈阳CR1-沈阳CR2': '沈阳CR1-北京CR1-北京CR2-沈阳CR2', '南京CR1-南京CR2': '南京CR1-上海CR1-上海CR2-南京CR2', '武汉CR1-武汉CR2': '武汉CR1-广州CR1-广州CR2-武汉CR2', '成都CR1-成都CR2': '成都CR1-武汉CR1-武汉CR2-成都CR2', '西安CR1-西安CR2': '西安CR1-武汉CR1-武汉CR2-西安CR2', '杭州CR1-杭州CR2': '杭州CR1-上海CR1-上海CR2-杭州CR2', '哈尔滨CR1-哈尔滨CR2': '哈尔滨CR1-沈阳CR1-沈阳CR2-哈尔滨CR2', '郑州CR1-郑州CR2': '郑州CR1-武汉CR1-武汉CR2-郑州CR2'}
        # print(te_pro) #{'北京-上海': '北京-（武汉）-上海', '北京-广州': '北京-（上海）-广州', '北京-沈阳': '北京-（哈尔滨）-沈阳', '北京-武汉': '北京-（西安）-武汉', '北京-西安': '北京-（武汉）-西安', '上海-广州': '上海-（武汉）-广州', '上海-南京': '上海-（杭州）-南京', '上海-武汉': '上海-（广州）-武汉', '上海-杭州': '上海-（广州）-杭州', '广州-武汉': '广州-（上海）-武汉', '广州-成都': '广州-（武汉）-（西安）-成都'}

        wentai_book = xlrd.open_workbook(os.path.join(os.path.abspath('excel/ipwin'), 'new_book.xls'))
        wentai_sh = wentai_book.sheet_by_name('中继转发流量')
        nodes = wentai_sh.row_values(0)[1:]

        wentai_nrow = wentai_sh.nrows
        work_book1 = xlwt.Workbook(encoding='utf-8')
        sheet1 = work_book1.add_sheet('稳态流量', cell_overwrite_ok=True)
        dict = {}
        for i in range(wentai_nrow):
            for j in range(wentai_nrow):
                val = wentai_sh.cell(i, j).value
                sheet1.write(i, j, val)

            sheet1.write(i, 0, d1[nodes[i - 1]])
            sheet1.write(0, i, d1[nodes[i - 1]])

        for key in te_CR.keys():

            sh = work_book1.add_sheet(key, cell_overwrite_ok=True)
            for i in range(1, wentai_nrow):
                sh.write(i, 0, d1[nodes[i - 1]])
                sh.write(0, i, d1[nodes[i - 1]])
                for j in range(1, wentai_nrow):
                    val = wentai_sh.cell(i, j).value
                    sh.write(i, j, val)
                    if val != 0 and val != '' and ('CR' in d1[nodes[j - 1]] and 'CR' in d1[nodes[i - 1]]):
                        dict[d1[nodes[j - 1]] + '-' + d1[nodes[i - 1]]] = val

            # TE 主备路由计算
            old_val = key
            new_val = te_CR[key]

            old = old_val.split('-')
            new = new_val.split('-')
            # 通过value找对应的key
            x = list(d1.keys())[list(d1.values()).index(old[1])]
            y = list(d1.keys())[list(d1.values()).index(old[0])]
            flow = wentai_sh.cell(nodes.index(x) + 1, nodes.index(y) + 1).value
            if flow == '':
                flow = 0
                sh.write(nodes.index(x) + 1, nodes.index(y) + 1, '', te_style)
            else:
                sh.write(nodes.index(x) + 1, nodes.index(y) + 1, '', te_style)

            for str_l in range(1, len(new)):
                startnode = new[str_l - 1]
                endnode = new[str_l]

                if (startnode + '-' + endnode) in list(dict.keys()):
                    # print(new)
                    # print(1, startnode + '-' + endnode)
                    x1 = list(d1.keys())[list(d1.values()).index(startnode)]
                    y1 = list(d1.keys())[list(d1.values()).index(endnode)]

                    row = nodes.index(x1) + 1
                    col = nodes.index(y1) + 1

                    flow2 = wentai_sh.cell(row, col).value
                    if flow2 == '' or flow2 == 0:

                        sh.write(col, row, 0 + flow, te_style2)
                    else:
                        sh.write(row, col, wentai_sh.cell(row, col).value + flow, te_style2)

        for key in te_pro.keys():

            sh = work_book1.add_sheet(key, cell_overwrite_ok=True)
            for i in range(1, wentai_nrow):
                sh.write(i, 0, d1[nodes[i - 1]])
                sh.write(0, i, d1[nodes[i - 1]])
                for j in range(1, wentai_nrow):
                    val = wentai_sh.cell(i, j).value
                    sh.write(i, j, val)

            # TE 主备路由计算
            old_val = key
            new_val = te_pro[key]

            new = new_val.replace('（', '').replace('）', '').replace('(', '').replace(')', '').split('-')

            for i in list(dict.keys()):
                n = re.sub("[A-Za-z0-9\/]", "", i)
                if old_val == n:

                    old = i.split('-')
                    x = list(d1.keys())[list(d1.values()).index(old[1])]
                    y = list(d1.keys())[list(d1.values()).index(old[0])]
                    flow = wentai_sh.cell(nodes.index(x) + 1, nodes.index(y) + 1).value
                    if flow == '':
                        flow = 0
                    sh.write(nodes.index(x) + 1, nodes.index(y) + 1, '', te_style)

                    for str_l in range(1, len(new)):

                        startnode = new[str_l - 1]
                        endnode = new[str_l]

                        if (startnode + '-' + endnode) == n:
                            newn = i.split('-')

                            x1 = list(d1.keys())[list(d1.values()).index(newn[1])]
                            y1 = list(d1.keys())[list(d1.values()).index(newn[0])]

                            flow2 = wentai_sh.cell(nodes.index(x1) + 1, nodes.index(y1) + 1).value
                            if flow2 == '':
                                flow2 = 0

                            sh.write(nodes.index(x1) + 1, nodes.index(y1) + 1, flow2 + flow, te_style2)

        p = os.path.join(os.path.abspath('excel/ipwin'), 'TE最大流量估算.xls')
        if os.path.exists(p):
            try:
                os.remove(p)
                work_book1.save(p)
            except PermissionError as e:
                logger.error(e)
                print(1, e)
                return JsonResponse({'status': 'fail', 'msg': '同名文件打开！！！'})
            except IOError as e:
                print(2, e)
                logger.error(e)
                return JsonResponse({'status': 'fail', 'msg': '没有检测到本期工程计算！！！'})
        else:
            work_book1.save(p)

        return JsonResponse({'status': 'ok', 'msg': '计算完成'})
    else:
        return JsonResponse({'status': 'fail', 'msg': '计算失败'})

def TE_max():
    p = os.path.join(os.path.abspath('excel/ipwin'), 'TE最大流量估算.xls')
    wb = xlrd.open_workbook(p,formatting_info=True)
    sheets = wb.sheet_names()
    fro = getMaxFiro(sheets)
    wb2 = copy(wb)
    ws = wb2.add_sheet('TE最大流量估算',cell_overwrite_ok=True)
    nrow = wb.sheet_by_index(0).nrows
    ncol = wb.sheet_by_index(0).ncols
    for i in range(1,nrow):
        for j in range(1,ncol):
            val = wb.sheet_by_index(0).cell(i,0).value
            ws.write(i,0,val)
            ws.write(0,i,val)
            ws.write(i,j,getFiro2(fro,j,i))
    wb2.save(p)

def ip_gz(request):
    if request.method == 'GET':
        pro_type = request.GET.get('pro_type')
        path1 = os.path.abspath('excel/ipwin/max/')
        if not os.path.exists(path1):
            os.mkdir(path1)

        node_dict = find_node()
        node_list = list(node_dict.keys())

        wb = xlrd.open_workbook(os.path.join(os.path.abspath('excel/ipwin'), 'IP网_直联.xls'))
        sheet1 = wb.sheet_by_index(0)
        num = len(node_list)

        # link = []

        dict = {}
        for i in range(1, num + 1):
            for j in range(1, num + 1):

                ban = sheet1.cell(i, j).value

                if ban != '' and ban != 0:
                    dict[node_list[i - 1] + '-' + node_list[j - 1]] = ban

                # l = []
                # l.append(node_list[i - 1])
                # l.append(node_list[j - 1])
                # link.append(l)
        # print(dict) #{'北C1-北C2': 608.0, '北C1-上C1': 1569.0, '北C1-广C1': 2387.0, '北C1-浙C1': 2075.0, 。。。
        key_list = list(dict.keys())
        # print(key_list) #['北C1-北C2', '北C1-上C1', '北C1-广C1', '北C1-浙C1',

        CR_node = [i for i in node_dict if 'C' in i]
        # print(CR_node)#['北C1', '北C2', '上C1', '上C2', '广C1', '广C2', '江C1', '江C2', '浙C1', '浙C2', '湖C1', '湖C2', '四C1', '四C2', '陕C1', '陕C2', '辽C1', '辽C2', '河C1', '河C2', '黑C1', '黑C2']

        max_fn(key_list, pro_type)
        lhx_fn(CR_node, pro_type)

        all_exce = glob.glob(path1 + '/' + "*.xls")
        # 得到要合并的所有exce表格数据
        if (all_exce == 0):
            logger.info("该目录下无.xlsx文件！请检查您输入的目录是否有误！")
            os.system('pause')
            exit()

        for exce in all_exce:

            file = os.path.basename(exce)

            if len(file) < 8:

                na = file[:-4]

                fh = open_exce(exce)
                sheet = fh.sheet_by_name('直联')  # 打开sheet页
                sheet1 = fh.sheet_by_name('点点路由')
                sheet_list = fh.sheet_names()
                nodes = sheet.row_values(0)[1:]
                remove_list = ['源节点', '目的节点', '合计']
                for i in remove_list:
                    if i in nodes:
                        nodes.remove(i)
                num = len(nodes)

                book2 = copy(fh)

                sheet_index = sheet_list.index('点点路由')
                sh = book2.get_sheet(sheet_index)

                nlist = []

                for i in range(1, num + 1):
                    elist = []
                    for j in range(1, num + 1):

                        bandwidth = sheet.cell_value(i, j)

                        if i == j or bandwidth == '' or bandwidth == 0:
                            continue
                        else:
                            edge = Edge(nodes[i - 1], nodes[j - 1], bandwidth)

                            if nodes[i - 1] == edge.startNodeId:
                                elist.append(edge)
                    if elist == []:
                        continue
                    else:
                        node = Node(nodes[i - 1], elist)
                    nlist.append(node)

                graph = Graph(nlist)

                for i in nodes:

                    startNodeId = i

                    for j in nodes:
                        endNodeId = j
                        if startNodeId == na or endNodeId == na:
                            continue

                        else:

                            originNodeId, destNodeId, route, weight = graph.dijkstra(startNodeId, endNodeId)

                        for j in range(1, num + 1):
                            for z in range(1, num + 1):
                                StoE = sheet1.cell((j - 1) * num + z, 0).value
                                if originNodeId == destNodeId:
                                    sh.write((j - 1) * num + z, 1, 'NULL')
                                    sh.write((j - 1) * num + z, 2, 'NULL')
                                    sh.write((j - 1) * num + z, 3, 'NULL')
                                    sh.write((j - 1) * num + z, 4, 'NULL')

                                if StoE == originNodeId + "-" + destNodeId:
                                    sh.write((j - 1) * num + z, 1, len(route))
                                    sh.write((j - 1) * num + z, 2, weight)

                                    for p in route:
                                        pa = p + [destNodeId]
                                        sh.write((j - 1) * num + z, 3, len(pa) - 1)
                                        sh.write((j - 1) * num + z, 4 + route.index(p), '-'.join(pa))
                                logger.info(originNodeId, destNodeId, '计算中。。。。。')
            else:
                fh = open_exce(exce)
                sheet = fh.sheet_by_name('直联')  # 打开sheet页
                sheet1 = fh.sheet_by_name('点点路由')
                sheet_list = fh.sheet_names()
                nodes = sheet.row_values(0)[1:]
                remove_list = ['源节点', '目的节点', '合计']
                for i in remove_list:
                    if i in nodes:
                        nodes.remove(i)
                num = len(nodes)

                book2 = copy(fh)

                sheet_index = sheet_list.index('点点路由')
                sh = book2.get_sheet(sheet_index)

                nlist = []

                for i in range(1, num + 1):
                    elist = []
                    for j in range(1, num + 1):

                        bandwidth = sheet.cell_value(i, j)

                        if i == j or bandwidth == '' or bandwidth == 0:
                            continue
                        else:
                            edge = Edge(nodes[i - 1], nodes[j - 1], bandwidth)

                            if nodes[i - 1] == edge.startNodeId:
                                elist.append(edge)
                    if elist == []:
                        continue
                    else:
                        node = Node(nodes[i - 1], elist)
                    nlist.append(node)

                graph = Graph(nlist)

                for i in nodes:

                    startNodeId = i

                    for j in nodes:
                        endNodeId = j
                        if startNodeId == endNodeId:
                            continue

                        else:

                            originNodeId, destNodeId, route, weight = graph.dijkstra(startNodeId, endNodeId)

                        for j in range(1, num + 1):
                            for z in range(1, num + 1):
                                StoE = sheet1.cell((j - 1) * num + z, 0).value
                                if originNodeId == destNodeId:
                                    sh.write((j - 1) * num + z, 1, 'NULL')
                                    sh.write((j - 1) * num + z, 2, 'NULL')
                                    sh.write((j - 1) * num + z, 3, 'NULL')
                                    sh.write((j - 1) * num + z, 4, 'NULL')

                                if StoE == originNodeId + "-" + destNodeId:
                                    sh.write((j - 1) * num + z, 1, len(route))
                                    sh.write((j - 1) * num + z, 2, weight)

                                    for p in route:
                                        pa = p + [destNodeId]
                                        sh.write((j - 1) * num + z, 3, len(pa) - 1)
                                        sh.write((j - 1) * num + z, 4 + route.index(p), '-'.join(pa))
                                logger.info(originNodeId, destNodeId, '计算中。。。。。')

            if os.path.exists(exce):
                book2.save(os.path.join(path1, 'book_new.xls'))  # 保存新的excel，保存excel必须使用后缀名是.xls的，不是能是.xlsx的
                os.remove(exce)
                os.rename(os.path.join(path1, 'book_new.xls'), exce)
            else:
                book2.save(exce)

            # 带宽分配
            bandwidth_allocation(pro_type, exce, os.path.join(path1, 'book_new.xls'))
            band_gz_max()

        return JsonResponse({'status': 'ok', 'msg': '计算完成'})
    else:
        return JsonResponse({'status': 'fail', 'msg': '计算失败'})


# 搜索故障文件
def search_routefile(request):
    if request.method == 'POST':
        node_key = request.POST.get('val')

        dir_name = os.path.abspath('excel/ipwin/max')
        file_name = os.listdir(dir_name)  # 路径下文件名称
        route_list = []
        for i in file_name:
            route_list.append(i[:-4])

        res = []
        for i in route_list:
            if node_key in i:
                res.append(i)
        jsonArr = json.dumps(res, ensure_ascii=False)

        return JsonResponse({'status': 'ok', 'msg': jsonArr})


# 故障计算后，导出点点路由表
def down_route(request):
    if request.method == 'POST':
        route = request.POST.get('node1')

        path = os.path.join(os.path.abspath('excel/ipwin/max'), route + '.xls')

        wb = xlrd.open_workbook(path, formatting_info=True)
        ws = wb.sheet_by_name('点点路由')
        book = xlwt.Workbook()
        sh = book.add_sheet('点点路由', cell_overwrite_ok=True)
        row = ws.nrows
        col = ws.ncols
        for i in range(row):
            for j in range(col):
                rou = ws.cell(i, j).value

                sh.write(i, j, rou)

        response = HttpResponse(content_type='application/octet-stream')
        response['Content-Disposition'] = 'attachment; filename="{0}"'.format(escape_uri_path('IP网_' + route + '.xls'))
        book.save(response)

        return response


# 输出流量取最大
def band_gz_max():
    path = os.path.abspath('excel/ipwin')

    work_book = xlwt.Workbook()
    sheet = work_book.add_sheet('输出带宽最大值', cell_overwrite_ok=True)

    # 新建表表头
    wb = xlrd.open_workbook(path + 'new_book.xls')
    sh = wb.sheet_by_name('带宽')
    nodes = sh.row_values(0)[1:]
    r = sh.nrows

    for i in range(1, r):
        sheet.write(i, 0, nodes[i - 1])
        sheet.write(0, i, nodes[i - 1])

    work_book.save(os.path.join(path, '输出带宽最大值.xls'))

    path1 = os.path.abspath('excel/ipwin/max/')
    if not os.path.exists(path1):
        return JsonResponse({'status': 'fail', 'msg': '计算失败未检测到故障表,请进行故障轮训计算'})

    all_exce = glob.glob(path1 + '/' + "*.xls")
    # 得到要合并的所有exce表格数据
    if (all_exce == 0):
        logger.info("该目录下无.xlsx文件！请检查您输入的目录是否有误！")
        os.system('pause')
        exit()

    p1 = os.path.join(path, '输出带宽最大值.xls')
    p2 = os.path.join(path, '输出带宽max.xls')

    for exc in all_exce:
        max_f = xlrd.open_workbook(p1, formatting_info=True)
        max_sh = max_f.sheet_by_index(0)
        newf = copy(max_f)
        newsh = newf.get_sheet(0)
        f = open_exce(exc)
        sh = f.sheet_by_name('中继转发流量')
        for i in range(1, r):
            for j in range(1, r):
                v = sh.cell(i, j).value
                v2 = max_sh.cell(i, j).value
                if v == '':
                    v = 0
                elif v2 == '':
                    v2 = 0

                if v > v2:
                    newsh.write(i, j, v)
                else:
                    continue

        if os.path.exists(p1):
            newf.save(p2)  # 保存新的excel，保存excel必须使用后缀名是.xls的，不是能是.xlsx的
            os.remove(p1)
            os.rename(p2, p1)
            logger.info('190+输出流量最大筛选完成')
        else:
            newf.save(p1)
            logger.info('190+输出流量最大筛选完成')


# 中继计算（10GE，100GE，400GE）
def iprelay_calculation(request):
    if request.method == 'POST':
        link_use = request.POST.get('link_use')
        v1 = request.POST.get('v1')
        v2 = request.POST.get('v2')
        granularity = request.POST.get('granularity')
        xs = ProjectInfo.objects.get(id=1).xs
        val = int(link_use) / xs

        path1 = os.path.join(os.path.abspath('excel/ipwin'), 'new_book.xls')
        path2 = os.path.join(os.path.abspath('excel/ipwin'), 'book_new.xls')

        book = xlrd.open_workbook(path1, formatting_info=True)  # 打开文件，并且保留原格式
        sheet_list = book.sheet_names()
        book2 = copy(book)
        sheet = book.sheet_by_name('中继转发流量')  # 打开sheet页

        ndict = find_node()
        nodes = sheet.row_values(0)[1:]
        if granularity == '100GE':
            if '中继带宽100GE' in sheet_list:
                sheet_index = sheet_list.index('中继带宽100GE')
                sh2 = book2.add_sheet('中继带宽折算', cell_overwrite_ok=True)
                sh = book2.get_sheet(sheet_index)
                for i in range(1, len(nodes) + 1):
                    c = sheet.cell(i, 0).value
                    sh.write(i, 0, ndict[c])
                    sh.write(0, i, ndict[c])
                    sh2.write(i, 0, ndict[c])
                    sh2.write(0, i, ndict[c])
                    for j in range(1, len(nodes) + 1):
                        flow = sheet.cell_value(i, j)
                        if flow != '':
                            T1 = math.ceil(flow / val / (100 * 1000))
                            sh.write(i, j, T1)
                            if T1 >= v2:
                                T2 = math.ceil(T1/v2) #向上取整
                                sh2.write(i, j, str(T2)+'*400GE')
                            else:
                                sh2.write(i, j, str(T1)+'*100GE')

            else:
                worksheet = book2.add_sheet('中继带宽100GE', cell_overwrite_ok=True)
                worksheet2 = book2.add_sheet('中继带宽折算', cell_overwrite_ok=True)

                for i in range(1, len(nodes) + 1):

                    c = sheet.cell(i, 0).value
                    worksheet.write(i, 0, ndict[c])
                    worksheet.write(0, i, ndict[c])
                    worksheet2.write(i, 0, ndict[c])
                    worksheet2.write(0, i, ndict[c])

                    for j in range(1, len(nodes) + 1):
                        flow = sheet.cell_value(i, j)
                        if flow != '':
                            T1 = math.ceil(flow / val / (100 * 1000))
                            worksheet.write(i, j, T1)
                            if T1 >= v2:
                                T2 = math.ceil(T1 / v2)  # 向上取整
                                worksheet2.write(i, j, str(T2) + '*400GE')
                            else:
                                worksheet2.write(i, j, str(T1) + '*100GE')

        elif granularity == '10GE':
            if '中继带宽10GE' in sheet_list:
                sheet_index = sheet_list.index('中继带宽10GE')
                sh2 = book2.add_sheet('中继带宽折算', cell_overwrite_ok=True)
                sh = book2.get_sheet(sheet_index)
                for i in range(1, len(nodes) + 1):
                    c = sheet.cell(i, 0).value
                    sh.write(i, 0, ndict[c])
                    sh.write(0, i, ndict[c])
                    sh2.write(i, 0, ndict[c])
                    sh2.write(0, i, ndict[c])
                    for j in range(1, len(nodes) + 1):
                        flow = sheet.cell_value(i, j)
                        if flow != '':
                            # T1 = round(flow/bandwidth_availability/10,0)*10
                            T1 = math.ceil(flow / val / (10 * 1000))
                            sh.write(i, j, T1)
                            if T1 >= v1:
                                T2 = math.ceil(T1/v1) #向上取整
                                sh2.write(i, j, str(T2)+'*100GE')
                            else:
                                sh2.write(i, j, str(T1)+'*10GE')


            else:
                worksheet = book2.add_sheet('中继带宽10GE', cell_overwrite_ok=True)
                worksheet2 = book2.add_sheet('中继带宽折算', cell_overwrite_ok=True)
                for i in range(len(nodes) + 1):
                    c = sheet.cell(i, 0).value
                    worksheet.write(i, 0, ndict[c])
                    worksheet.write(0, i, ndict[c])
                    worksheet2.write(i, 0, ndict[c])
                    worksheet2.write(0, i, ndict[c])

                for i in range(1, len(nodes) + 1):
                    for j in range(1, len(nodes) + 1):
                        flow = sheet.cell_value(i, j)
                        if flow != '':
                            T1 = math.ceil(flow / val / (10 * 1000))

                            worksheet.write(i, j, T1)
                            if T1 >= v2:
                                T2 = math.ceil(T1 / v1)  # 向上取整
                                worksheet2.write(i, j, str(T2) + '*100GE')
                            else:
                                worksheet2.write(i, j, str(T1) + '*10GE')

        elif granularity == '400GE':
            if '中继带宽400GE' in sheet_list:
                sheet_index = sheet_list.index('中继带宽400GE')
                sh = book2.get_sheet(sheet_index)
                for i in range(1, len(nodes) + 1):
                    c = sheet.cell(i, 0).value
                    sh.write(i, 0, ndict[c])
                    sh.write(0, i, ndict[c])
                    for j in range(1, len(nodes) + 1):
                        flow = sheet.cell_value(i, j)
                        if flow != '':
                            # T1 = round(flow/bandwidth_availability/10,0)*10
                            T3 = math.ceil(flow / val / (400 * 1000))
                            sh.write(i, j, T3)



            else:
                worksheet = book2.add_sheet('中继带宽400GE', cell_overwrite_ok=True)

                for i in range(1, len(nodes) + 1):
                    c = sheet.cell(i, 0).value
                    worksheet.write(i, 0, ndict[c])
                    worksheet.write(0, i, ndict[c])
                    for j in range(1, len(nodes) + 1):
                        flow = sheet.cell_value(i, j)
                        if flow != '':
                            # T1 = round(flow/bandwidth_availability/10,0)*10
                            T3 = math.ceil(flow / val / (400 * 1000))
                            worksheet.write(i, j, T3)

        if os.path.exists(path1):
            os.remove(path1)
            book2.save(path2)  # 保存新的excel，保存excel必须使用后缀名是.xls的，不是能是.xlsx的

            os.rename(path2, path1)
        else:
            book2.save(path1)
        return JsonResponse({'status': 'ok', 'msg': '计算完成'})
    else:
        return JsonResponse({'status': 'ok', 'msg': '计算失败'})


# 中继比较
def iprelay_compare(request):
    if request.method == 'POST':
        relayFile = request.FILES.get('relayfile')

        if not (relayFile):
            # 有数据为空
            return JsonResponse({'status': 'err', 'msg': '参数不能为空!'})

        path = settings.MEDIA_ROOT + 'ipwin/relay/'  # 上传文件的保存路径，可以自己指定任意的路径
        path1 = os.path.join(os.path.abspath('excel/ipwin'), 'new_book.xls')
        path2 = os.path.join(os.path.abspath('excel/ipwin'), 'book_new.xls')
        wb = xlrd.open_workbook(path1, formatting_info=True)
        book2 = copy(wb)
        sheets = wb.sheet_names()

        if not os.path.exists(path):
            os.makedirs(path)

        with open(os.path.join(path, relayFile.name), 'wb+') as f:
            for chunk in relayFile.chunks():
                f.write(chunk)
        try:

            # ================== 往期数据提取 =============================================

            excel = xlrd.open_workbook(path + relayFile.name)  # 往期
            sheet = excel.sheet_by_index(0)
            nodes = sheet.row_values(0)[1:]

            remove_list = ['最终中继链路需求', '', '合计']
            for i in remove_list:
                if i in nodes:
                    nodes.remove(i)
            oldnrows = sheet.nrows
            oldncols = sheet.ncols

            olddict = {}
            for i in range(3, oldnrows):
                for j in range(1, oldncols):
                    relay = sheet.cell(i, j).value
                    if relay != 0 and relay != '':
                        olddict[nodes[i - 3] + '~' + nodes[j - 1]] = relay

            # ================== 本期数据提取 =============================================

            wb = xlrd.open_workbook(path1, formatting_info=True)  # 本期
            if '100' in relayFile.name:
                different_province_circuit(path1, path2)  # 初始化分省路由表（总）
                ws = wb.sheet_by_name('中继带宽100GE')


            elif '10GE' in relayFile.name:
                different_province_circuit(path1, path2)  # 初始化分省路由表（总）
                ws = wb.sheet_by_name('中继带宽10GE')

            newnodes = ws.row_values(0)[1:]
            nrows = ws.nrows
            ncols = ws.ncols

            dict = {}
            for i in range(1, nrows):
                for j in range(1, ncols):
                    relay = ws.cell(i, j).value
                    if relay != 0 and relay != '':
                        dict[newnodes[i - 1] + '~' + newnodes[j - 1]] = relay

            new_link = list(set(list(olddict.keys()) + list(dict.keys())))

            index = sheets.index('分省电路表')
            sh = book2.get_sheet(index)

            for i in range(2, len(new_link) + 1):
                for j in range(2, 15):
                    if i != len(new_link):
                        sh.write(i + 2, 0, i - 1, style2)
                        sh.write(i + 2, j - 1, '', style2)
                    else:

                        sh.write(i + 2, 0, '', he_style3)
                        sh.write(i + 2, 1, '合计', he_style3)
                        sh.write(i + 2, 2, Formula('SUM(C3:C62)'), he_style3)
                        sh.write(i + 2, 3, Formula('SUM(D3:D62)'), he_style3)
                        sh.write(i + 2, 4, Formula('SUM(E3:E62)'), he_style3)
                        sh.write(i + 2, 5, Formula('SUM(F3:F62)'), he_style3)
                        sh.write(i + 2, 6, Formula('SUM(G3:G62)'), he_style3)
                        sh.write(i + 2, 7, Formula('SUM(H3:H62)'), he_style3)
                        sh.write(i + 2, 8, Formula('SUM(I3:I62)'), he_style3)
                        sh.write(i + 2, 9, Formula('SUM(J3:J62)'), he_style3)
                        sh.write(i + 2, 10, Formula('SUM(K3:K62)'), he_style3)
                        sh.write(i + 2, 11, Formula('SUM(L3:L62)'), he_style3)
                        sh.write(i + 2, 12, Formula('SUM(M3:M62)'), he_style3)
                        sh.write(i + 2, 13, Formula('SUM(N3:N62)'), he_style3)

                for k, v in olddict.items():
                    # print(k, v)
                    sh.write(new_link.index(k) + 2, 1, k, style3)
                    if isinstance(v, float):
                        sh.write(new_link.index(k) + 2, 2, v, style3)
                    else:
                        n = v.split('*')

                        if 'POS' in v:
                            sh.write(new_link.index(k) + 2, 3, int(n[0]), style3)
                        elif '100GE' in v:
                            sh.write(new_link.index(k) + 2, 2, int(n[0]), style3)
                        elif '10GE' in v:
                            sh.write(new_link.index(k) + 2, 4, int(n[0]), style3)
                        elif isinstance(v, float):
                            sh.write(new_link.index(k) + 2, 2, int(n[0]), style3)

                for k, v in dict.items():
                    # print(k, v)
                    sh.write(new_link.index(k) + 2, 1, k, style3)
                    if isinstance(v, float):
                        sh.write(new_link.index(k) + 2, 2, v, style3)
                    else:
                        if '100' in relayFile.name:
                            sh.write(new_link.index(k) + 2, 5, v, style3)
                        elif '10GE' in relayFile.name:
                            sh.write(new_link.index(k) + 2, 7, v, style3)
            if '400' in relayFile.name:
                different_province_circuit400(path1, path2)  # 初始化分省路由表（总）
                ws = wb.sheet_by_name('中继带宽400GE')

                newnodes = ws.row_values(0)[1:]
                nrows = ws.nrows
                ncols = ws.ncols

                dict = {}
                for i in range(1, nrows):
                    for j in range(1, ncols):
                        relay = ws.cell(i, j).value
                        if relay != 0 and relay != '':
                            dict[newnodes[i - 1] + '~' + newnodes[j - 1]] = relay

                new_link = list(set(list(olddict.keys()) + list(dict.keys())))

                index = sheets.index('分省电路表')
                sh = book2.get_sheet(index)

                for i in range(2, len(new_link) + 1):
                    sh.write(i + 2, 0, '', he_style3)
                    sh.write(i + 2, 1, '合计', he_style3)
                    sh.write(i + 2, 2, Formula('SUM(C3:C' + str(len(new_link)) + ')'), he_style3)
                    sh.write(i + 2, 3, Formula('SUM(D3:D' + str(len(new_link)) + ')'), he_style3)
                    sh.write(i + 2, 4, Formula('SUM(E3:E' + str(len(new_link)) + ')'), he_style3)
                    sh.write(i + 2, 5, Formula('SUM(F3:F' + str(len(new_link)) + ')'), he_style3)
                    sh.write(i + 2, 6, Formula('SUM(G3:G' + str(len(new_link)) + ')'), he_style3)
                    sh.write(i + 2, 7, Formula('SUM(H3:H' + str(len(new_link)) + ')'), he_style3)
                    sh.write(i + 2, 8, Formula('SUM(I3:I' + str(len(new_link)) + ')'), he_style3)
                    sh.write(i + 2, 9, Formula('SUM(J3:J' + str(len(new_link)) + ')'), he_style3)
                    sh.write(i + 2, 10, Formula('SUM(K3:K' + str(len(new_link)) + ')'), he_style3)
                    sh.write(i + 2, 11, Formula('SUM(L3:L' + str(len(new_link)) + ')'), he_style3)
                    sh.write(i + 2, 12, Formula('SUM(M3:M' + str(len(new_link)) + ')'), he_style3)
                    sh.write(i + 2, 13, Formula('SUM(N3:N' + str(len(new_link)) + ')'), he_style3)
                    sh.write(i + 2, 14, Formula('SUM(O3:O' + str(len(new_link)) + ')'), he_style3)
                    sh.write(i + 2, 15, Formula('SUM(P3:P' + str(len(new_link)) + ')'), he_style3)
                    sh.write(i + 2, 16, Formula('SUM(Q3:Q' + str(len(new_link)) + ')'), he_style3)
                    sh.write(i + 2, 17, Formula('SUM(R3:R' + str(len(new_link)) + ')'), he_style3)
                    for j in range(2, 19):
                        if i != len(new_link):
                            sh.write(i + 2, 0, i - 1, style2)
                            sh.write(i + 2, j - 1, '', style2)

                for k, v in olddict.items():
                    # print(k, v)
                    sh.write(new_link.index(k) + 2, 1, k, style3)
                    if isinstance(v, float):
                        sh.write(new_link.index(k) + 2, 2, v, style3)
                    else:
                        n = v.split('*')

                        if 'POS' in v:
                            sh.write(new_link.index(k) + 2, 3, int(n[0]), style3)
                        elif '100GE' in v:
                            sh.write(new_link.index(k) + 2, 2, int(n[0]), style3)
                        elif '10GE' in v:
                            sh.write(new_link.index(k) + 2, 4, int(n[0]), style3)
                        elif '400GE' in v:
                            sh.write(new_link.index(k) + 2, 5, int(n[0]), style3)
                        elif isinstance(v, float):
                            sh.write(new_link.index(k) + 2, 2, int(n[0]), style3)

                for k, v in dict.items():
                    # print(k, v)
                    sh.write(new_link.index(k) + 2, 1, k, style3)
                    if isinstance(v, float):
                        sh.write(new_link.index(k) + 2, 2, v, style3)
                    else:
                        if '100' in relayFile.name:
                            sh.write(new_link.index(k) + 2, 6, v, style3)
                        elif '10GE' in relayFile.name:
                            sh.write(new_link.index(k) + 2, 8, v, style3)
                        elif '400GE' in v:
                            sh.write(new_link.index(k) + 2, 9, v, style3)
                        elif isinstance(v, float):
                            sh.write(new_link.index(k) + 2, 9, v, style3)

            if os.path.exists(path1):
                book2.save(path2)
                os.remove(path1)

                os.rename(path2, path1)
            else:
                book2.save(path1)

            logger.info('往期数据整合计算完成！！')

            b = xlrd.open_workbook(path1, formatting_info=True)
            b_ws = b.sheet_by_name('分省电路表')
            cols = b_ws.ncols
            b2 = copy(b)
            b2_ws_index = b.sheet_names().index('分省电路表')
            b2_ws = b2.get_sheet(b2_ws_index)

            if cols > 15:
                for i in range(2, len(new_link) + 1):

                    if b_ws.cell(i, 2) - b_ws.cell(i, 6) < 0:
                        b2_ws.write(i, 10, 0)
                        b2_ws.write(i, 114, -(b_ws.cell(i, 2) - b_ws.cell(i, 6)))

                    else:
                        b2_ws.write(i, 10, b_ws.cell(i, 2) - b_ws.cell(i, 5))
                        b2_ws.write(i, 14, 0)

                    if b_ws.cell(i, 3) - b_ws.cell(i, 7) < 0:
                        b2_ws.write(i, 11, 0)
                        b2_ws.write(i, 15, -(b_ws.cell(i, 3) - b_ws.cell(i, 6)))

                    else:
                        b2_ws.write(i, 11, b_ws.cell(i, 2) - b_ws.cell(i, 5))
                        b2_ws.write(i, 15, 0)

                    if b_ws.cell(i, 4) - b_ws.cell(i, 8) < 0:
                        b2_ws.write(i, 12, 0)
                        b2_ws.write(i, 16, -(b_ws.cell(i, 4) - b_ws.cell(i, 7)))

                    else:
                        b2_ws.write(i, 12, b_ws.cell(i, 2) - b_ws.cell(i, 5))
                        b2_ws.write(i, 16, 0)

                    if b_ws.cell(i, 5) - b_ws.cell(i, 9) < 0:
                        b2_ws.write(i, 13, 0)
                        b2_ws.write(i, 17, -(b_ws.cell(i, 4) - b_ws.cell(i, 7)))

                    else:
                        b2_ws.write(i, 13, b_ws.cell(i, 2) - b_ws.cell(i, 5))
                        b2_ws.write(i, 17, 0)

            else:
                for i in range(2, len(new_link) + 1):

                    if b_ws.cell(i, 2) - b_ws.cell(i, 5) < 0:
                        b2_ws.write(i, 8, 0)
                        b2_ws.write(i, 11, -(b_ws.cell(i, 2) - b_ws.cell(i, 5)))

                    else:
                        b2_ws.write(i, 8, b_ws.cell(i, 2) - b_ws.cell(i, 5))
                        b2_ws.write(i, 11, 0)

                    if b_ws.cell(i, 3) - b_ws.cell(i, 6) < 0:
                        b2_ws.write(i, 9, 0)
                        b2_ws.write(i, 12, -(b_ws.cell(i, 3) - b_ws.cell(i, 6)))

                    else:
                        b2_ws.write(i, 9, b_ws.cell(i, 2) - b_ws.cell(i, 5))
                        b2_ws.write(i, 12, 0)

                    if b_ws.cell(i, 4) - b_ws.cell(i, 7) < 0:
                        b2_ws.write(i, 10, 0)
                        b2_ws.write(i, 13, -(b_ws.cell(i, 4) - b_ws.cell(i, 7)))

                    else:
                        b2_ws.write(i, 10, b_ws.cell(i, 2) - b_ws.cell(i, 5))
                        b2_ws.write(i, 13, 0)

            if os.path.exists(path1):
                b2.save(path2)
                os.remove(path1)

                os.rename(path2, path1)
            else:
                b2.save(path1)

            logger.info('新增撤销整合计算完成！！')

            return JsonResponse({'status': 'ok', 'msg': '计算完成！！'})
        except PermissionError as e:
            logger.error(e)
            return JsonResponse({'status': 'err', 'msg': '你有可能已经打开了这个文件，关闭这个文件即可！！'})

        except IOError as e:
            logger.error(e)

            return JsonResponse({'status': 'err', 'msg': '没有检测到本期工程计算，请先进行计算！！'})

    else:
        return JsonResponse({'status': 'err', 'msg': '计算失败，请重新计算！！'})


def report_down_load(request):
    if request.method == 'POST':
        filename = request.POST.get('filename')

        path = os.path.abspath('excel/ipwin')

        wb = xlrd.open_workbook(os.path.join(path, 'new_book.xls'), formatting_info=True)
        wb2 = copy(wb)

        if not filename:
            pro_name = ProjectInfo.objects.filter(id=1).values('name')#<QuerySet [{'name': 'IP专网骨干网十一期工程'}]> <class 'django.db.models.query.QuerySet'>
            pro_name = list(pro_name)[0]['name']

            response = HttpResponse(content_type='application/octet-stream')
            response['Content-Disposition'] = 'attachment; filename="{0}"'.format(
                escape_uri_path(pro_name + '中继仿真.xls'))
            wb2.save(response)
            return response
        else:

            response = HttpResponse(content_type='application/vnd.ms-excel')
            response['Content-Disposition'] = 'attachment; filename=' + filename + '.xls'
            wb2.save(response)
            return response
