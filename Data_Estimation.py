#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Sep  9 17:23:49 2020

@author: ziling
"""

import xlrd
import xlwt
import scipy.stats as stats
from scipy import stats
import numpy as np

file = input("Please input your file path for stage2:")
by_name = input("please input your sheet name in that file:")
# file = 'Data_set.xlsx'
# by_name = 'Sheet1'
def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception:
        print ('e')


def set_style(name,height,bold=False):
	style = xlwt.XFStyle()
	font = xlwt.Font()
	font.name = name
	font.bold = bold
	font.color_index = 4
	font.height = height
	style.font = font
	return style

def excel_table_byname():
    data = open_excel(file) #Open excel
    table = data.sheet_by_name(by_name)#obtain the sheet in Excel file by name
    book = xlwt.Workbook() #create Excel file
    sheet1 = book.add_sheet('sheet1')

    col0=table.col_values(0)
    for i in range(0,len(col0)):
        sheet1.write(i,0,str(col0[i]))
        book.save('Data_estimation.xls')
    col1 = table.col_values(1)

    for j in range(0,len(col1)):
        sheet1.write(j,1,col1[j])
        book.save('Data_estimation.xls')
    
#find random number generation range    
    max_value = table.col_values(-5) 
    max_value.pop(0)
    est_max = table.col_values(-1)
    est_max.pop(0)    
    for n in range(0,len(max_value)):
        if max_value[n] >= est_max[n]:
            max_value[n] = max_value[n]
        if max_value[n] < est_max[n]:
            max_value[n] = est_max[n] #maximum range
    stddev = table.col_values(-4)
    stddev.pop(0)
    for i in range(0,len(stddev)):
        if stddev[i] == 0:
            stddev[i] = 0.000001
    
    mean_value = table.col_values(-6)
    mean_value.pop(0)
    min_value = table.col_values(-7)
    min_value.pop(0)
    est_min = table.col_values(-3)
    est_min.pop(0)
    for m in range(0,len(min_value)):
        if min_value[m] >= est_min[m]:
            min_value[m] = est_min[m]
        if min_value[m] < est_min[m]:
            min_value[m] = min_value[m] #minimum range

#random integer generation based on the min and max range
    data_est = []
    for a in range(0,len(max_value)):
        # dist = stats.truncnorm((min_value[a]-mean_value[a])/stddev[a], (max_value[a]-mean_value[a])/stddev[a], loc = mean_value[a], scale = stddev[a])
        dist = stats.norm(loc=mean_value[a],scale=stddev[a])
        data = dist.rvs(size=30)

        # dist = np.random.normal(loc=mean_value[a],scale=stddev[a],size=30)
        # data = dist.rvs(30) #as the collecting sample size is small and less than 30, 
#using t-distribution, the maximum sample size for t-distribution is 30
        data_est.append(data)
    data_est = np.vstack(data_est)
    # print(data_est)
    [h, l] = data_est.shape
    for b in range(h):
        min_ = min_value[b]
        for c in range(l):
            m = float(data_est[b,c])
            if m < min_:
                m=min_
            sheet1.write(b+1, c+2, m)
            book.save('Data_estimation.xls')
    
    for i in range(30):
        sheet1.write(0,i+2,"estimated_set %d"%(int(i+1)))
        book.save('Data_estimation.xls')
    
    return 'Data_estimation.xls'
        

def input_file():
        return [file, by_name]

