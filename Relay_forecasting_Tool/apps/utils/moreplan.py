import os
import xlrd
import xlwt
from xlutils.copy import copy
import glob


# 自定义表格样式
style = xlwt.easyxf('align: wrap on, vert centre, horiz center;'
                    'borders:left 1,right 1,top 1,bottom 1')

#方案对比
def option_comparison(path):
    dir_list = os.listdir(path)
    file_list = []  # 该列表用来存放当前目录下的所有txt文件

    print('当前文件夹下的所有excel文件:')


    for l in dir_list:

        if l.rfind('xls') >= 0:

            print('             ', l)


        file_list.append(l)
    all_exce = glob.glob(path +"/*.xls")
    print(all_exce)

    print(dir_list)
    print(file_list)


path = 'D:\job\Relay_forecasting_Tool\excel\cmnetwin'
option_comparison(path)