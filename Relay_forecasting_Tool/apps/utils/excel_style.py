from xlutils.copy import copy
import xlrd
from xlwt import *
import os


# 分省电路
def different_province_circuit(path1,path2):

    book = xlrd.open_workbook(path1,formatting_info=True)
    sheet_list = book.sheet_names()
    new_book = copy(book)
    if '分省电路表' in sheet_list:
        index = sheet_list.index('分省电路表')
        sheet = new_book.get_sheet(index)
    else:
        sheet = new_book.add_sheet('分省电路表',cell_overwrite_ok=True)

    sec_col = sheet.col(1)
    sec_col.width = 256 * 20

    style = XFStyle() #赋值style为XFStyle()，初始化样式
    fnt = Font()  # 创建一个文本格式，包括字体、字号和颜色样式特性
    fnt.name = u'宋体'  # 设置其字体为微软雅黑
    fnt.height = 260
    fnt.bold = True
    style.font = fnt  #将赋值好的模式参数导入Style

    #设置居中
    al = Alignment()
    al.horz = 0x02      # 设置水平居中
    al.vert = 0x01      # 设置垂直居中
    style.alignment = al

    borders = Borders() # 设置单元格下框线样式
    borders.left = 2
    borders.right = 2
    borders.top = 2
    borders.bottom = 2
    style.borders = borders  # 将赋值好的模式参数导入Style

    style1 = XFStyle() #赋值style为XFStyle()，初始化样式
    fnt = Font()  # 创建一个文本格式，包括字体、字号和颜色样式特性
    fnt.name = u'Calibri'  # 设置其字体为宋体
    fnt.height = 260
    style1.font = fnt  #将赋值好的模式参数导入Style

    #设置居中
    al = Alignment()
    al.horz = 0x02      # 设置水平居中
    al.vert = 0x01      # 设置垂直居中
    style1.alignment = al

    borders = Borders() # 设置单元格下框线样式
    borders.left = 2
    borders.right = 2
    borders.top = 2
    borders.bottom = 2
    style1.borders = borders  # 将赋值好的模式参数导入Style

    style2 = XFStyle() #赋值style为XFStyle()，初始化样式

    fnt = Font()  # 创建一个文本格式，包括字体、字号和颜色样式特性
    fnt.name = u'Calibri'  # 设置其字体为宋体
    fnt.height = 260
    fnt.bold = True
    style2.font = fnt  #将赋值好的模式参数导入Style

    #设置居中
    al = Alignment()
    al.horz = 0x02      # 设置水平居中
    al.vert = 0x01      # 设置垂直居中
    style2.alignment = al

    borders = Borders() # 设置单元格下框线样式
    borders.left = 2
    borders.right = 2
    borders.top = 2
    borders.bottom = 2
    style2.borders = borders  # 将赋值好的模式参数导入Style


    Line_data1 = (u'序号') # 创建一个Line_data列表，并将其值赋为测试表，以utf-8编码时中文前加u
    Line_data2 = (u'电路局向')
    if 'cmnet' in path1:

        Line_data3 = (u'CMNET往期到达')
        Line_data4 = (u'CMNET本期到达')
    elif 'ip' in path1:
        Line_data3 = (u'IP往期到达')
        Line_data4 = (u'IP本期到达')
    Line_data5 = (u'本期新增')
    Line_data6 = (u'本期撤销')

    '''
    write_merge(start_row,end_row,start_col,end_col,data,style)中的6个参数
    
        start_row：合并单元格的起始行  
        end_row：合并单元格的终止行    
        start_col：合并单元格的起始列  
        end_col：合并单元格的终止列  
        data：内容   
        style：样式
    '''

    sheet.write_merge(0, 1, 0, 0, Line_data1, style) #以合并单元格形式写入数据，即将数据写入
    sheet.write_merge(0, 1, 1, 1, Line_data2, style)
    sheet.write_merge(0, 0, 2, 4, Line_data3, style)
    sheet.write_merge(1, 1, 2, 2, u'100GE', style1)
    sheet.write_merge(1, 1, 3, 3, u'10GPOS', style1)
    sheet.write_merge(1, 1, 4, 4, u'10GE', style1)
    sheet.write_merge(0, 0, 5, 7, Line_data4, style)
    sheet.write_merge(1, 1, 5, 5, u'100GE', style1)
    sheet.write_merge(1, 1, 6, 6, u'10GPOS', style1)
    sheet.write_merge(1, 1, 7, 7, u'10GE', style1)
    sheet.write_merge(0, 0, 8, 10, Line_data5, style)
    sheet.write_merge(1, 1, 8, 8, u'100GE', style1)
    sheet.write_merge(1, 1, 9, 9, u'10GPOS', style1)
    sheet.write_merge(1, 1, 10, 10, u'10GE', style1)
    sheet.write_merge(0, 0, 11, 13, Line_data6, style)
    sheet.write_merge(1, 1, 11, 11, u'100GE', style1)
    sheet.write_merge(1, 1, 12, 12, u'10GPOS', style1)
    sheet.write_merge(1, 1, 13, 13, u'10GE', style1)


    if os.path.exists(path1):
        new_book.save(path2)
        os.remove(path1)

        os.rename(path2, path1)
    else:
        new_book.save(path1)  # 保存新的excel，保存excel必须使用后缀名是.xls的，不是能是.xlsx的


# 分省电路
def different_province_circuit400(path1, path2):
    book = xlrd.open_workbook(path1, formatting_info=True)
    sheet_list = book.sheet_names()
    new_book = copy(book)
    if '分省电路表' in sheet_list:
        index = sheet_list.index('分省电路表')
        sheet = new_book.get_sheet(index)
    else:
        sheet = new_book.add_sheet('分省电路表', cell_overwrite_ok=True)

    sec_col = sheet.col(1)
    sec_col.width = 256 * 20

    style = XFStyle()  # 赋值style为XFStyle()，初始化样式
    fnt = Font()  # 创建一个文本格式，包括字体、字号和颜色样式特性
    fnt.name = u'宋体'  # 设置其字体为微软雅黑
    fnt.height = 260
    fnt.bold = True
    style.font = fnt  # 将赋值好的模式参数导入Style

    # 设置居中
    al = Alignment()
    al.horz = 0x02  # 设置水平居中
    al.vert = 0x01  # 设置垂直居中
    style.alignment = al

    borders = Borders()  # 设置单元格下框线样式
    borders.left = 2
    borders.right = 2
    borders.top = 2
    borders.bottom = 2
    style.borders = borders  # 将赋值好的模式参数导入Style
    style1 = XFStyle()  # 赋值style为XFStyle()，初始化样式

    fnt = Font()  # 创建一个文本格式，包括字体、字号和颜色样式特性
    fnt.name = u'Calibri'  # 设置其字体为宋体
    fnt.height = 260
    style1.font = fnt  # 将赋值好的模式参数导入Style

    # 设置居中
    al = Alignment()
    al.horz = 0x02  # 设置水平居中
    al.vert = 0x01  # 设置垂直居中
    style1.alignment = al

    borders = Borders()  # 设置单元格下框线样式
    borders.left = 2
    borders.right = 2
    borders.top = 2
    borders.bottom = 2
    style1.borders = borders  # 将赋值好的模式参数导入Style

    style2 = XFStyle()  # 赋值style为XFStyle()，初始化样式

    fnt = Font()  # 创建一个文本格式，包括字体、字号和颜色样式特性
    fnt.name = u'Calibri'  # 设置其字体为宋体
    fnt.height = 260
    fnt.bold = True
    style2.font = fnt  # 将赋值好的模式参数导入Style

    # 设置居中
    al = Alignment()
    al.horz = 0x02  # 设置水平居中
    al.vert = 0x01  # 设置垂直居中
    style2.alignment = al

    borders = Borders()  # 设置单元格下框线样式
    borders.left = 2
    borders.right = 2
    borders.top = 2
    borders.bottom = 2
    style2.borders = borders  # 将赋值好的模式参数导入Style

    Line_data1 = (u'序号')  # 创建一个Line_data列表，并将其值赋为测试表，以utf-8编码时中文前加u
    Line_data2 = (u'电路局向')
    if 'CMNET' in path1:

        Line_data3 = (u'CMNET往期到达')
        Line_data4 = (u'CMNET本期到达')
        sheet.write_merge(0, 0, 2, 5, Line_data3, style)
        sheet.write_merge(0, 0, 6, 9, Line_data4, style)
    elif 'IP' in path1:
        Line_data3 = (u'IP往期到达')
        Line_data4 = (u'IP本期到达')
        sheet.write_merge(0, 0, 2, 5, Line_data3, style)
        sheet.write_merge(0, 0, 6, 9, Line_data4, style)
    Line_data5 = (u'本期新增')
    Line_data6 = (u'本期撤销')

    '''
    write_merge(start_row,end_row,start_col,end_col,data,style)中的6个参数

        start_row：合并单元格的起始行  
        end_row：合并单元格的终止行    
        start_col：合并单元格的起始列  
        end_col：合并单元格的终止列  
        data：内容   
        style：样式
    '''

    sheet.write_merge(0, 1, 0, 0, Line_data1, style)  # 以合并单元格形式写入数据，即将数据写入
    sheet.write_merge(0, 1, 1, 1, Line_data2, style)
    sheet.write_merge(0, 0, 2, 5, Line_data3, style)
    sheet.write_merge(0, 0, 6, 9, Line_data4, style)
    sheet.write_merge(0, 0, 10, 13, Line_data5, style)
    sheet.write_merge(0, 0, 14, 17, Line_data6, style)

    #=====================往期到达==============================

    sheet.write_merge(1, 1, 2, 2, u'100GE', style1)
    sheet.write_merge(1, 1, 3, 3, u'10GPOS', style1)
    sheet.write_merge(1, 1, 4, 4, u'10GE', style1)
    sheet.write_merge(1, 1, 5, 5, u'400GE', style1)

    # =====================本期到达==============================

    sheet.write_merge(1, 1, 6, 6, u'100GE', style1)
    sheet.write_merge(1, 1, 7, 7, u'10GPOS', style1)
    sheet.write_merge(1, 1, 8,8, u'10GE', style1)
    sheet.write_merge(1, 1,9, 9, u'400GE', style1)

    # =====================本期新增到达==============================

    sheet.write_merge(1, 1, 10, 10, u'100GE', style1)
    sheet.write_merge(1, 1, 11, 11, u'10GPOS', style1)
    sheet.write_merge(1, 1, 12, 12, u'10GE', style1)
    sheet.write_merge(1, 1, 13, 13, u'400GE', style1)

    # =====================本期撤销到达==============================

    sheet.write_merge(1, 1, 14, 14, u'100GE', style1)
    sheet.write_merge(1, 1, 15, 15, u'10GPOS', style1)
    sheet.write_merge(1, 1, 16, 16, u'10GE', style1)
    sheet.write_merge(1, 1, 17, 17, u'400GE', style1)


    if os.path.exists(path1):
        new_book.save(path2)
        os.remove(path1)

        os.rename(path2, path1)
    else:
        new_book.save(path1)  # 保存新的excel，保存excel必须使用后缀名是.xls的，不是能是.xlsx的



