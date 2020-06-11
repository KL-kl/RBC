from django.http import JsonResponse
import xlrd
import xlwt
from xlutils.copy import copy
import os
import glob
from utils.hebin import *
from Relay_forecasting_Tool.settings import MEDIA_ROOT

# 故障轮询 ---- 类核心断
style = xlwt.easyxf('pattern: pattern solid, fore_colour red; ')


def lhx_fn(nodelist,pro_type):
    if pro_type == '2':

        path = os.path.abspath('excel/cmnetwin/max/')
        if not os.path.exists(path):
            os.mkdir(path)

        path2 = os.path.join(os.path.abspath('excel/cmnetwin/'), 'CMNET网_metric.xls')

        if not os.path.exists(path2):
            return JsonResponse({'status': 'err', 'msg': '未找到metric表，请先导入'})
        wb = xlrd.open_workbook(path2, formatting_info=True)

        ws = wb.sheet_by_name('Metric')
        nodes = ws.row_values(0)[1:]

        rows = ws.nrows  # 63
        cols = ws.ncols  # 63
        li = ['' for r in range(1, rows)]
        book = xlwt.Workbook(encoding='utf-8')

        sh = book.add_sheet('Metric', cell_overwrite_ok=True)
        sh2 = book.add_sheet('直联', cell_overwrite_ok=True)
        sh3 = book.add_sheet('点点路由', cell_overwrite_ok=True)

        for k in nodelist:

            for i in range(1, rows):

                for j in range(1, cols):

                    ban = ws.cell(i, j).value
                    sh.write(0, j, nodes[j - 1])
                    sh.write(j, 0, nodes[j - 1])
                    sh2.write(0, j, nodes[j - 1])
                    sh2.write(j, 0, nodes[j - 1])

                    if ban == 0 or ban == '':
                        continue
                    else:
                        if i == (nodes.index(k) + 1) or j == (nodes.index(k) + 1):
                            sh2.write(nodes.index(k) + 1, j, li[j-1])
                            sh2.write(i, nodes.index(k) + 1, li[j-1])
                            sh.write(nodes.index(k) + 1, j, li[j-1])
                            sh.write(i, nodes.index(k) + 1, li[j-1])
                        else:
                            sh.write(i, j, ban)

                            sh2.write(i, j, ban)
                            sh2.write(j, i, ban)

                    start_node = ws.cell_value(i, 0)
                    end_node = ws.cell_value(0, j)

                    sh3.write((i - 1) * len(nodes) + j, 0, start_node + '-' + end_node)
                    if i == j:
                        sh3.write((i - 1) * len(nodes) + j, 1, 'NULL')
                        sh3.write((i - 1) * len(nodes) + j, 2, 'NULL')
                        sh3.write((i - 1) * len(nodes) + j, 3, 'NULL')
                        sh3.write((i - 1) * len(nodes) + j, 4, 'NULL')


            sh3.write(0, 0, '源节点-目的节点')
            sh3.write(0, 1, '同权路由数量')
            sh3.write(0, 2, '度量值')
            sh3.write(0, 3, '路由跳数')
            sh3.write(0, 4, '路由')

            book.save(os.path.join(path, k + '.xls'))

    elif pro_type == '1':
        path = os.path.abspath('excel/ipwin/max/')
        if not os.path.exists(path):
            os.mkdir(path)

        path2 = os.path.join(MEDIA_ROOT, r'ipwin/metric')
        path2 = path2.replace('\\', '/')
        file = 'IP网_metric.xls'

        if not os.path.exists(os.path.join(path2, file)):
            return JsonResponse({'status': 'err', 'msg': '未找到metric表，请先导入'})
        wb = xlrd.open_workbook(os.path.join(path2, file), formatting_info=True)

        ws = wb.sheet_by_name('Metric')
        nodes = ws.row_values(0)[1:]

        rows = ws.nrows  # 63
        cols = ws.ncols  # 63
        li = ['' for r in range(1, rows)]
        wb3 = copy(wb)

        sh = wb3.get_sheet(0)
        sh2 = wb3.add_sheet('直联', cell_overwrite_ok=True)
        sh3 = wb3.add_sheet('点点路由', cell_overwrite_ok=True)

        for k in nodelist:

            for i in range(1, rows):
                sh.write(nodes.index(k) + 1, i, '')
                sh.write(i, nodes.index(k) + 1, '')
                for j in range(1, cols):

                    ban = ws.cell(i, j).value
                    sh2.write(0, j, nodes[j - 1])
                    sh2.write(j, 0, nodes[j - 1])

                    if ban == 0 or ban == '':
                        ban = ''
                    if i == (nodes.index(k) + 1):
                        sh2.write(nodes.index(k) + 1, j, '')

                    elif j == (nodes.index(k) + 1):
                        sh2.write(i, nodes.index(k) + 1, li[j])
                    else:
                        sh2.write(i, j, ban)
                        sh2.write(j, i, ban)

                    start_node = ws.cell_value(i, 0)
                    end_node = ws.cell_value(0, j)

                    sh3.write((i - 1) * len(nodes) + j, 0, start_node + '-' + end_node)
                    if i == j:
                        sh3.write((i - 1) * len(nodes) + j, 1, 'NULL')
                        sh3.write((i - 1) * len(nodes) + j, 2, 'NULL')
                        sh3.write((i - 1) * len(nodes) + j, 3, 'NULL')
                        sh3.write((i - 1) * len(nodes) + j, 4, 'NULL')

            sh3.write(0, 0, '源节点-目的节点')
            sh3.write(0, 1, '同权路由数量')
            sh3.write(0, 2, '度量值')
            sh3.write(0, 3, '路由跳数')
            sh3.write(0, 4, '路由')

            wb3.save(os.path.join(path, k + '.xls'))

