#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Sep  9 17:23:53 2020

@author: ziling
"""

import sympy
import pandas as pd
import  xdrlib ,sys
import Data_Estimation
import xlrd
import xlwt


# def data_estimation():
#     file = Data_Estimation.excel_table_byname()
#     return 'Data_estimation.xls'

def open_excel(file= 'Data_estimation.xls'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception:
        print ('e')

#get data from excel based on the name   parameters:file：Excel path     colnameindex：index of column name  ，by_name：name of Sheet1
def set_style(name,height,bold=False):
	style = xlwt.XFStyle()
	font = xlwt.Font()
	font.name = name
	font.bold = bold
	font.color_index = 4
	font.height = height
	style.font = font
	return style

# def feasible_dataset():
#     time_max = xlrd.open_workbook('time_range.xls').sheet_by_name('sheet1').col_values(2)[1:]
#     return time_max
    
def excel_table(file= 'Data_estimation.xls', colnameindex=0, by_name=u'sheet1'):
    # de = data_estimation()
    # time_max = feasible_dataset()
    data = open_excel(file) ##open excel
    table = data.sheet_by_name(by_name)#get sheet from excel file based on sheet name
   # nrows = table.nrows #行数 
    stakeholders = table.col_values(1)[1:] #data of one col
    book = xlwt.Workbook() #create an Excel
    sheet1 = book.add_sheet('new_dataset') #creat a sheet
    i = 0 #index of row
    row0=table.row_values(0)
    row0.pop(1)
    raw_data = xlrd.open_workbook('data_set.xlsx').sheet_by_name('Sheet1')
    row00=raw_data.row_values(0)
    row00.pop(1)
    row0=row0+row00[-7:]
    row2=[]
    	#write first row
    for i in range(0,len(row0)):
        sheet1.write(0,i,str(row0[i]))
        book.save('new_dataset.xls')
    t = 1
    c=1
    y=0

    
    for j in stakeholders:
        if j == 1.0:
            row=table.row_values(t)
            row.pop(1)
            row1 = row[0:len(row)]
            row2 = raw_data.row_values(t)
            # for m in row1:
            #     if type(m) is float:
            #         if m <= time_max[y]:
            #             row1 = [' ' if x == m else x for x in row1]
            # y = y+1
            for i in range(0,len(row1)+7):
                row3=row1+row2[-7:]
                sheet1.write(c,i,str(row3[i]),set_style('Times New Roman',220,True))
                book.save('new_dataset.xls')
            c = c+1
        t = t+1
    return 'new_dataset.xls'

# if __name__ =="__main__":
#   excel_table(file= 'Data_estimation.xls', colnameindex=0, by_name=u'sheet1')


