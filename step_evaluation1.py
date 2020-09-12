#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Sep  9 17:23:53 2020

@author: ziling
"""

import simpy
from simpy.events import AllOf
import pandas as pd
import  xdrlib ,sys
import xlrd
import xlwt
file = 'Data_set.xlsx'

def open_excel(file= 'Data_set.xlsx'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception:
        print (str(e))

#根据名称获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的索引  ，by_name：Sheet1名称
def set_style(name,height,bold=False):
	style = xlwt.XFStyle()
	font = xlwt.Font()
	font.name = name
	font.bold = bold
	font.color_index = 4
	font.height = height
	style.font = font
	return style

def feasible_dataset():
    time_max = xlrd.open_workbook('time_range.xlsx').sheet_by_name('Sheet1').col_values(2)[1:]
    return time_max
    
def excel_table_byname(file= 'Data_set.xlsx', colnameindex=0, by_name=u'Sheet1'):
    time_max = feasible_dataset()
    data = open_excel(file) #打开excel文件
    table = data.sheet_by_name(by_name) #根据sheet名字来获取excel中的sheet
   # nrows = table.nrows #行数 
    stakeholders = table.col_values(1)[1:] #某一列数据 
    book = xlwt.Workbook() #创建一个Excel
    sheet1 = book.add_sheet('new_dataset') #在其中创建一个名为hello的sheet
    i = 0 #行序号
    row0=table.row_values(0)
    row0.pop(1)
    	#写第一行
    for i in range(0,len(row0)):
        sheet1.write(0,i,str(row0[i]))
        book.save('new_dataset.xls')
    t = 1
    for j in stakeholders:
        if j == 1.0:
            row=table.row_values(t)
            row.pop(1)
            row1 = row[0:len(row)-3]
            y = 0
            for m in row1:
                if type(m) is float:
                    if m <= time_max[y]:
                        row1 = [' ' if x == m else x for x in row1]
                        row2 = row1+row[-3:]
                y = y+1
            for i in range(0,len(row2)):
                sheet1.write(t,i,str(row2[i]),set_style('Times New Roman',220,True))
                book.save('new_dataset.xls')
            t = t+1
    return 'new_dataset.xls'

if __name__ =="__main__":
  excel_table_byname(file= 'Data_set.xlsx', colnameindex=0, by_name=u'Sheet1')

