from django.http import JsonResponse
import xlrd
import xlwt
from xlutils.copy import copy
import os
import glob
from Relay_forecasting_Tool.settings import MEDIA_ROOT
#故障轮询 ---- 全网断

style = xlwt.easyxf('pattern: pattern solid, fore_colour red; ')


def max_fn(key_list,pro_type):
    if pro_type == '2':

        path = os.path.abspath('excel/cmnetwin/max/')

        if not os.path.exists(path):
            os.mkdir(path)

        path2 = os.path.join(MEDIA_ROOT, r'cmnetwin/metric')
        path2 = path2.replace('\\', '/')
        file = 'CMNET网_metric.xls'

        if not os.path.exists(os.path.join(path2, file)):
            return JsonResponse({'status': 'err', 'msg': '未找到metric表，请先导入'})

        wb = xlrd.open_workbook(os.path.join(path2, file), formatting_info=True)
        ws = wb.sheet_by_name('Metric')
        nodes = ws.row_values(0)[1:]
        cols = ws.ncols

        for k in key_list:
            wb3 = copy(wb)
            sh2 = wb3.add_sheet('直联', cell_overwrite_ok=True)
            sh3 = wb3.add_sheet('点点路由', cell_overwrite_ok=True)

            node = k.split('-')
            sh = wb3.get_sheet(0)
            sh.write(nodes.index(node[0]) + 1, nodes.index(node[1]) + 1, '', style)

            for i in range(1, cols):
                start_node = ws.cell(i, 0).value
                for j in range(1, cols):
                    end_node = ws.cell(0, j).value
                    sh2.write(i, 0, start_node)
                    sh2.write(0, j, end_node)
                    sh3.write((i - 1) * len(nodes) + j, 0, start_node + '-' + end_node)
                    if i == j:
                        sh3.write((i - 1) * len(nodes) + j, 1, 'NULL')
                        sh3.write((i - 1) * len(nodes) + j, 2, 'NULL')
                        sh3.write((i - 1) * len(nodes) + j, 3, 'NULL')
                        sh3.write((i - 1) * len(nodes) + j, 4, 'NULL')

                    ban = ws.cell(i, j).value

                    if ban == 0 or ban == '':
                        ban = ''
                    if (i != (nodes.index(node[0]) + 1) and j != (nodes.index(node[1]) + 1)) or (
                            i != (nodes.index(node[1]) + 1) and j != (nodes.index(node[0]) + 1)):
                        sh2.write(i, j, ban)
                        sh2.write(j, i, ban)
                    else:
                        sh2.write(nodes.index(node[0]) + 1, nodes.index(node[1]) + 1, '', style)
                        sh2.write(nodes.index(node[1]) + 1, nodes.index(node[0]) + 1, '', style)

            sh3.write(0, 0, '源节点-目的节点')
            sh3.write(0, 1, '同权路由数量')
            sh3.write(0, 2, '度量值')
            sh3.write(0, 3, '路由跳数')
            sh3.write(0, 4, '路由')
            wb3.save(os.path.join(path, k + '.xls'))

    elif pro_type == '1':

        path = os.path.abspath('excel/ipwin/max/')

        if not os.path.exists(path):
            os.mkdir(path)

        # path2 = os.path.join(MEDIA_ROOT, r'ipwin/metric')
        # path2 = path2.replace('\\', '/')
        # file = 'IP网_metric.xls'
        #
        # if not os.path.exists(os.path.join(path2, file)):
        #     return JsonResponse({'status': 'err', 'msg': '未找到metric表，请先导入'})

        path2 = os.path.join(os.path.abspath('excel/ipwin'),'new_book.xls')


        wb = xlrd.open_workbook(path2, formatting_info=True)
        ws = wb.sheet_by_name('直联')
        nodes = ws.row_values(0)[1:]
        cols = ws.ncols

        for k in key_list:
            wb3 = xlwt.Workbook(encoding='utf-8')
            sh2 = wb3.add_sheet('直联', cell_overwrite_ok=True)
            sh3 = wb3.add_sheet('点点路由', cell_overwrite_ok=True)

            node = k.split('-')
            # sh = wb3.get_sheet(0)
            # sh.write(nodes.index(node[1]) + 1, nodes.index(node[0]) + 1, '',style)

            for i in range(1,cols):
                start_node = ws.cell(i, 0).value
                for j in range(1,cols):
                    end_node = ws.cell(0, j).value
                    sh2.write(i, 0, start_node)
                    sh2.write(0,j,end_node)

                    sh3.write((i - 1) * len(nodes) + j, 0, start_node + '-' + end_node)
                    if i == j:
                        sh3.write((i - 1) * len(nodes) + j, 1, 'NULL')
                        sh3.write((i - 1) * len(nodes) + j, 2, 'NULL')
                        sh3.write((i - 1) * len(nodes) + j, 3, 'NULL')
                        sh3.write((i - 1) * len(nodes) + j, 4, 'NULL')

                    ban = ws.cell(i, j).value

                    if ban == 0 or ban == '':

                        ban = ''
                    if (i != (nodes.index(node[0])+1) and j != (nodes.index(node[1])+1)) or (i != (nodes.index(node[1])+1) and j != (nodes.index(node[0])+1)):
                        sh2.write(i, j, ban)
                        sh2.write(j, i, ban)
                    else:
                        sh2.write(nodes.index(node[0])+1, nodes.index(node[1])+1, '', style)
                        sh2.write(nodes.index(node[1]) + 1, nodes.index(node[0]) + 1, '', style)

            sh3.write(0, 0, '源节点-目的节点')
            sh3.write(0, 1, '同权路由数量')
            sh3.write(0, 2, '度量值')
            sh3.write(0, 3, '路由跳数')
            sh3.write(0, 4, '路由')
            wb3.save(os.path.join(path,k+'.xls'))


