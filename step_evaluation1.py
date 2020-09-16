#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Sep  9 17:23:53 2020

@author: ziling
"""

import sympy
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

def feasible_dataset():
    time_max = xlrd.open_workbook('time_range.xlsx').sheet_by_name('Sheet1').col_values(2)[1:]
    return time_max
    
def excel_table_byname(file= 'Data_set.xlsx', colnameindex=0, by_name=u'Sheet1'):
    time_max = feasible_dataset()
    data = open_excel(file) #open excel
    table = data.sheet_by_name(by_name) #get sheet from excel file based on sheet name
   # nrows = table.nrows #number of rows
    stakeholders = table.col_values(1)[1:] #data of one col
    book = xlwt.Workbook() #create an Excel
    sheet1 = book.add_sheet('new_dataset') #creat a sheet
    i = 0 #index of row
    row0=table.row_values(0)
    row0.pop(1)
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
            row1 = row[0:len(row)-3]
            for m in row1:
                if type(m) is float:
                    if m <= time_max[y]:
                        row1 = [' ' if x == m else x for x in row1]
            y = y+1
            for i in range(0,len(row1)+3):
                row2=row1+row[-3:]
                print(row2)
                sheet1.write(c,i,str(row2[i]),set_style('Times New Roman',220,True))
                book.save('new_dataset.xls')
            c = c+1
        t = t+1
    return 'new_dataset.xls'

if __name__ =="__main__":
  excel_table_byname(file= 'Data_set.xlsx', colnameindex=0, by_name=u'Sheet1')


