#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Sep  9 17:26:56 2020

@author: ziling
"""

import sympy
import pandas as pd
import  xdrlib ,sys
import xlrd
import xlwt
import numpy as np
from scipy.stats import t
from numpy import average, std
from math import sqrt


file = 'Data_estimation.xls'
def open_excel(file= 'Data_estimation.xls'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception:
        print (str(e))

def set_style(name,height,bold=False):
	style = xlwt.XFStyle()
	font = xlwt.Font()
	font.name = name
	font.bold = bold
	font.color_index = 4
	font.height = height
	style.font = font
	return style

def excel_table_byname(file= 'Data_estimation.xls', colnameindex=0, by_name=u'sheet1'):
    data = open_excel(file) #Open excel
    table = data.sheet_by_name(by_name)#obtain the sheet in Excel file by name
    book = xlwt.Workbook() #create Excel file
    sheet1 = book.add_sheet('sheet1')
    
    col0=table.col_values(0)
    for i in range(0,len(col0)):
        sheet1.write(i,0,str(col0[i]))
        book.save('ideal_range.xls')
    
    row0 = ['lmin','lmax','hmin','hmax']
    for j in range(0,len(row0)):
        sheet1.write(0,j+1,str(row0[j]))
        book.save('ideal_range.xls')
#input the raw data as a matrix
    set_matrix=[]
    for row in range (1,table.nrows):
        _row = []
        for col in range (1,table.ncols-8):
            _row.append(table.cell_value(row,col+1))
        set_matrix.append(_row)
    set_matrix_array = np.array(set_matrix)

    #sample mean
    mean = table.col_values(-6)
    mean.pop(0)

    #sample variance (The degree used in calculations is N - ddof), the calculation is done in Excel
    stddev = table.col_values(-4)
    stddev.pop(0)

    [h,l] = set_matrix_array.shape
    #due to the small sample size, using t-distribution, the confidence level is 95%
    t_bounds = t.interval(0.9, l - 1,mean,stddev)
    t_bounds = np.vstack(t_bounds)
    t_bounds = t_bounds.transpose()
    [a, b] = t_bounds.shape
    for i in range(a):
        for j in range(b):
            sheet1.write(i+1, j+1, t_bounds[i, j])
            book.save('ideal_range.xls')
    large_bounds = t.interval(0.98, l - 1,mean,stddev)
    large_bounds = np.vstack(large_bounds)
    large_bounds = large_bounds.transpose()
    [c,d] = large_bounds.shape
    for i in range(c):
        for j in range(d):
            sheet1.write(i+1,j+3,large_bounds[i,j])
            book.save('ideal_range.xls')
        
        
        
        
# if __name__ =="__main__":
#   excel_table_byname(file= 'Data_estimation.xls', colnameindex=0, by_name=u'sheet1')