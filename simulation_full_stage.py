#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Sep  9 17:28:04 2020

@author: ziling
"""
import Data_Estimation
import numpy as np
import xlrd
import xlwt
from xlutils.copy import copy

file_sheet = Data_Estimation.input_file()
file = file_sheet[0]
sheet_name = file_sheet[1]
# full_stage_data=input('Please input your file path for full stage data:')
full_stage_data = 'full_stage_data.xlsx'
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

def old_data():
    data = open_excel(file) #Open excel
    data2 = open_excel('Data_estimation.xls')
    table = data.sheet_by_name(sheet_name)#obtain the sheet in Excel file by name
    table2 = data2.sheet_by_name('sheet1')
    time_set1 = xlrd.open_workbook('ideal_range.xls')
    time_set = copy(time_set1)
    #create new sheet to describe the type of steps
    old_data1 = time_set.add_sheet('stage2_old_data')
    new_data1 = time_set.add_sheet('stage2_new_data')
    old_data = time_set.get_sheet(2)    
    new_data = time_set.get_sheet(3)
    
    col0=table.col_values(0)
    for i in range(0,len(col0)):
        old_data.write(i,0,col0[i])
        time_set.save('ideal_range.xls')
    old_data.write(len(col0),0,'Total time spend')
    row0=table.row_values(0)
    for j in range(1,len(row0)-7):
        old_data.write(0,j,str(row0[j]))
        time_set.save('ideal_range.xls')
        
    row02=table2.row_values(0)
    row02.pop(0)
    row02.pop(0)
    for i in range(len(row02)):        
        old_data.write(0,i+len(row0)-7,str(row02[i]))
        time_set.save('ideal_range.xls')
        
    
    col1=table.col_values(1)
    col1.pop(0)
    for k in range(0,len(col1)):
        old_data.write(k+1,1,col1[k])
        time_set.save('ideal_range.xls')
        
    set_matrix=[]
    for row in range (1,table.nrows):
        _row = []
        for col in range (1,table.ncols-8):
            _row.append(table.cell_value(row,col+1))
        set_matrix.append(_row)
    set_matrix_array = np.array(set_matrix)
    total_time = np.sum(set_matrix,axis=0)
    total_set1 = np.vstack((set_matrix_array,total_time))
    
    set_matrix2=[]
    for row in range (1,table2.nrows):
        _row2 = []
        for col in range (2,table2.ncols):
            _row2.append(table2.cell_value(row,col))
        set_matrix2.append(_row2)
    set_matrix_array2 = np.array(set_matrix2)
    total_time2 = np.sum(set_matrix2,axis=0)
    total_set2 = np.vstack((set_matrix_array2,total_time2))
    
    total_set1t = total_set1.transpose()
    total_set2t = total_set2.transpose()
    total_sett = np.vstack((total_set1t,total_set2t))
    total_set = total_sett.transpose()
    
    [h, l] = total_set.shape
    for i in range(h):
        for j in range(l):
            old_data.write(i+1, j+2,total_set[i, j])
            time_set.save('ideal_range.xls')

def data_step():
    newdata = xlrd.open_workbook(file)
    step_ana = newdata.sheet_by_name(sheet_name)
    key = step_ana.col_values(0)[1:]
    n = 1
    dict_data = {}
    for i in key:
        value =[]
        for j in range(1,len(step_ana.row_values(0)[1:])-7):
            m = float(step_ana.row_values(n)[j+1])
            value.insert(j,m)
            dict_anadata = {i:value}
        dict_data.update(dict_anadata)
        n = n+1
    return dict_data

def optimization_stage2():
    dict_data=data_step()
    data = xlrd.open_workbook('ideal_range.xls')
    table = data.sheet_by_name('stage2_old_data')
    row0=table.row_values(0)
    steps = table.col_values(0)[1:]
    # book1 = xlrd.open_workbook('ideal_range.xls') #create an Excel
    book = copy(data)
    sheet1 = book.get_sheet(3) #creat a sheet
    time_set = xlrd.open_workbook('ideal_range.xls')
    time_range = time_set.sheet_by_name('sheet1')
    stage2_max = time_range.col_values(2)[1:]
    
    for i in range(0,len(row0)):
        sheet1.write(0,i,str(row0[i]))
        book.save('ideal_range.xls')
    
    for i in range(0,len(steps)-1):
        step = steps[i]
        sheet1.write(i+1,0,step)
        book.save('ideal_range.xls')
        step_max = stage2_max[i]
        if step in dict_data:
            stakeholder = table.col_values(1)[i+1]
            value = table.row_values(i+1)[1:]
            if stakeholder == 0:
                for j in range(0,len(value)):
                    sheet1.write(i+1,j+1,value[j])
                    book.save('ideal_range.xls')
            else:
                for j in range(0,len(value)):
                    step_d = value[j]
                    if step_d > step_max:
                        step_value = step_max
                        sheet1.write(i+1,j+1,step_value)
                        book.save('ideal_range.xls')
                    else:
                        sheet1.write(i+1,j+1,step_d)
                        book.save('ideal_range.xls')
    
def total_time():
    optimization_stage2()
    # dict_value = data_step()
    stage2_op1 = xlrd.open_workbook('ideal_range.xls')
    table1 = stage2_op1.sheet_by_name('stage2_new_data')
    steps_list = table1.col_values(0)
    stage2_op = copy(stage2_op1)
    table = stage2_op.get_sheet(3)
    table.write(len(steps_list),0,'Total Time Spend')
    stage2_op.save('ideal_range.xls')
    t = 1
    key = table1.row_values(0)[2:]
    n = 2
    dict_value = {}
    for i in key:
        value =[]
        for j in range(0,len(table1.col_values(0)[1:])):
            m = float(table1.col_values(n)[j+1])
            value.insert(j,m)
            dict_anadata = {i:value}
        dict_value.update(dict_anadata)
        n = n+1
    for keys in list(dict_value.keys()):
        value = dict_value[keys]
        sum_ = 0
        for i in value:
            sum_ = sum_+i
        table.write(len(steps_list),t+1,sum_)
        stage2_op.save('ideal_range.xls')
        t = t+1
        if t>len(dict_value.keys()):
            return
            
def average_time_reduced():
    # old = old_data()
    # total = total_time()
    stage2_op1 = xlrd.open_workbook('ideal_range.xls')
    new_data = stage2_op1.sheet_by_name('stage2_new_data')
    old_data = stage2_op1.sheet_by_name('stage2_old_data')
    new_sum = new_data.row_values(-1)[2:]
    old_sum = old_data.row_values(-1)[2:]
    reduce = []
    for i in range(0,len(new_sum)):
        old_time = old_sum[i]
        new_time = new_sum[i]
        # print(new_sum,new_time)
        diff = old_time - new_time
        reduce.append(diff)
    sum_ = 0
    for j in reduce:
        sum_ = sum_+j
    ave = sum_/len(reduce)
    stage2_op = copy(stage2_op1)
    output = stage2_op.get_sheet(1)
    output.write(len(stage2_op1.sheet_by_index(1).col_values(0)),0,'Average Time Reduce')
    output.write(len(stage2_op1.sheet_by_index(1).col_values(0)),1,ave)
    stage2_op.save('ideal_range.xls')

def full_stage_raw():
    full_stage = xlrd.open_workbook(full_stage_data)
    sheets = len(full_stage.sheets())
    output_file1 = xlrd.open_workbook('ideal_range.xls')
    output_file = copy(output_file1)
    for sheet in range(sheets):
        table = full_stage.sheet_by_index(sheet)
        rows = table.nrows
        cols = table.ncols
        worksheet = output_file.add_sheet('full_stage_raw')
        for i in range(0,rows):
            for j in range(0, cols):
                worksheet.write(i, j ,table.cell_value(i, j))
    output_file.save('ideal_range.xls')

def data_step1():
    full_stage_raw()
    newdata = xlrd.open_workbook('ideal_range.xls')
    step_ana = newdata.sheet_by_name('full_stage_raw')
    key = step_ana.col_values(1)[1:]
    n = 1
    dict_data = {}
    for i in key:
        value =[]
        for j in range(1,len(step_ana.row_values(0)[1:])):
            m = float(step_ana.row_values(n)[j+1])
            value.insert(j,m)
            dict_anadata = {i:value}
        dict_data.update(dict_anadata)
        n = n+1
    return dict_data

def total_time1():
    stage2_op1 = xlrd.open_workbook('ideal_range.xls')
    table1 = stage2_op1.sheet_by_name('full_stage_raw')
    steps_list = table1.col_values(1)
    data_n = table1.row_values(0)
    stage2_op = copy(stage2_op1)
    table = stage2_op.get_sheet(4)
    table.write(len(steps_list),0,'Total Time Spend')
    stage2_op.save('ideal_range.xls')
    t = 1
    key = table1.row_values(0)[2:]
    n = 2
    dict_value = {}
    for i in key:
        value =[]
        for j in range(0,len(table1.col_values(0)[1:])):
            m = float(table1.col_values(n)[j+1])
            value.insert(j,m)
            dict_anadata = {i:value}
        dict_value.update(dict_anadata)
        n = n+1
    for keys in list(dict_value.keys()):
        value = dict_value[keys]
        sum_ = 0
        for i in value:
            sum_ = sum_+i
        table.write(len(steps_list),t+1,sum_)
        stage2_op.save('ideal_range.xls')
        t = t+1
        if t>len(value):
            return
        
def optimization_stage2_full():
    dict_data=data_step()
    dict_fullstage = data_step1()
    # print(dict_fullstage)
    full_stage = xlrd.open_workbook(full_stage_data)
    sheets = len(full_stage.sheets())
    output_file1 = xlrd.open_workbook('ideal_range.xls')
    output_file = copy(output_file1)
    for sheet in range(sheets):
        table = full_stage.sheet_by_index(sheet)
        rows = table.nrows
        cols = table.ncols
        worksheet = output_file.add_sheet('full_stage_after_reduce')
        for i in range(0,rows):
            for j in range(0, cols):
                worksheet.write(i, j ,table.cell_value(i, j))
    output_file.save('ideal_range.xls')
    
    full_keys = list(dict_fullstage.keys())
    data = xlrd.open_workbook(file)
    table = data.sheet_by_name(sheet_name)
    steps = table.col_values(0)[1:]
    book1 = xlrd.open_workbook('ideal_range.xls') #create an Excel
    book = copy(book1)
    sheet1 = book.get_sheet(5)
    time_set = xlrd.open_workbook('ideal_range.xls')
    time_range = time_set.sheet_by_name('sheet1')
    stage2_max = time_range.col_values(2)[1:]
    for n in range(0,len(full_keys)):
        step_f = full_keys[n]
        for i in range(0,len(steps)):
            step = steps[i]
            step_max = stage2_max[i]
            if step_f == step:
                if step in dict_data:
                    stakeholder = table.col_values(1)[i+1]
                    value = table.row_values(i+1)[1:]
                    if stakeholder == 0:
                        for j in range(1,len(value)-7):
                            sheet1.write(n+1,j+1,value[j])
                            book.save('ideal_range.xls')
                    else:
                        for j in range(1,len(value)-7):
                            step_d = value[j]
                            if step_d > step_max:
                                step_value = step_max
                                sheet1.write(n+1,j+1,step_value)
                                book.save('ideal_range.xls')
                            else:
                                sheet1.write(n+1,j+1,step_d)
                                book.save('ideal_range.xls')

def total_time_after_reduce():
    old_data()
    total_time()
    average_time_reduced()
    optimization_stage2_full()
    total_time1()
    stage2_op1 = xlrd.open_workbook('ideal_range.xls')
    table1 = stage2_op1.sheet_by_name('full_stage_after_reduce')
    steps_list = table1.col_values(1)
    stage2_op = copy(stage2_op1)
    table = stage2_op.get_sheet(5)
    table.write(len(steps_list),0,'Total Time Spend')
    stage2_op.save('ideal_range.xls')
    t = 1
    key = table1.row_values(0)[2:]
    n = 2
    dict_value = {}
    for i in key:
        value =[]
        for j in range(0,len(table1.col_values(0)[1:])):
            m = float(table1.col_values(n)[j+1])
            value.insert(j,m)
            dict_anadata = {i:value}
        dict_value.update(dict_anadata)
        n = n+1
    for keys in list(dict_value.keys()):
        value = dict_value[keys]
        sum_ = 0
        for i in value:
            sum_ = sum_+i
        table.write(len(steps_list),t+1,sum_)
        stage2_op.save('ideal_range.xls')
        t = t+1
        if t>len(value):
            return
