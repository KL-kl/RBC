import os
from Relay_forecasting_Tool.settings import MEDIA_ROOT
import xlrd, csv
import xlwt
import glob
from xlutils.copy import copy
import numpy as np

biao_tou = "NULL"
wei_zhi = "NULL"

# 获取要合并的所有exce表格
def get_exce(pro_type):
    global wei_zhi
    # wei_zhi = input("请输入Exce文件所在的目录：")
    if pro_type == '2':
        wei_zhi = MEDIA_ROOT + r'cmnetwin/max/'
    elif pro_type == '1':
        wei_zhi = MEDIA_ROOT + r'ipwin/max/'

    all_exce = glob.glob(wei_zhi + "*.xls")
    print("该目录下有" + str(len(all_exce)) + "个exce文件：")
    if (len(all_exce) == 0):
        return 0
    else:
        # for i in range(len(all_exce)):
        #     print(all_exce[i])
        return all_exce

def open_exce(name):
    fh = xlrd.open_workbook(name,formatting_info=True)
    return fh

def add_sheet(exce,name):
    sh = exce.add_sheet(name,cell_overwrite_ok=True)
    return sh

def save_exce(newexce,oldpath,newpath):
    if os.path.exists(oldpath):
        newexce.save(newpath)
        os.remove(oldpath)

        os.rename(newpath, oldpath)
    else:
        newexce.save(oldpath)

def get_name(pro_type):

    if pro_type == '2':
        dir_name = MEDIA_ROOT + r'cmnetwin/max/'
        file_name = os.listdir(dir_name)  # 路径下文件名称
        file_dir = [os.path.join(dir_name, x) for x in file_name]  # 得到文件路径
        return file_name
    elif pro_type == '1':
        dir_name = MEDIA_ROOT + r'ipwin/max/'

        file_name = os.listdir(dir_name)  # 路径下文件名称
        file_dir = [os.path.join(dir_name, x) for x in file_name]  # 得到文件路径
        return file_name
