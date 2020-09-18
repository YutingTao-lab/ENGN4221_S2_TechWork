#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Sep 11 10:44:44 2020

@author: ziling
"""

import step_evaluation1
import four_loops_arrange
import sympy
# from sympy.events import AllOf
import pandas as pd
import  xdrlib ,sys
import xlrd
import xlwt

def not_feasible_step():
    #get the analyzable steps
    step_ana = xlrd.open_workbook(step_evaluation1.excel_table_byname(file= 'Data_set.xlsx', colnameindex=0,         by_name=u'Sheet1')).sheet_by_name('new_dataset')
    key = step_ana.col_values(0)[1:]
    n = 1
    dict_data = {}
    for i in key:
        #m = 1
        value =[]
        for j in range(0,len(step_ana.row_values(0)[1:])):
            value.append(step_ana.row_values(n)[j+1])
            #m = m+1
            for s in value:
                if s != ' ':
                    value = [float(s) if x == s else x for x in value]
                dict_anadata = {i:value}
        dict_data.update(dict_anadata)
        #print(dict_anadata)
        n = n+1
    return dict_data

def statistical_values():
    step_ana = xlrd.open_workbook('new_dataset.xls').sheet_by_name('new_dataset')
    key = step_ana.row_values(0)[-7:]
    n = 1
    dict_ref = {}
    for i in key:
        #m = 1
        value =[]
        for j in range(0,len(step_ana.col_values(0)[1:])):
            value.append(step_ana.col_values(-n)[j+1])
            #m = m+1
            for s in value:
                if s != ' ':
                    value = [float(s) if x == s else x for x in value]
                dict_statistical = {key[-n]:value}
        dict_ref.update(dict_statistical)
        n = n+1
    return dict_ref

def loop_and_steps():
    all_steps = xlrd.open_workbook('Data_set.xlsx').sheet_by_name('Sheet1').col_values(0)
    dict_loops={}
    # dict_loops = {'loopA':all_steps[0:2],'loopB':all_steps[2:9],'loopC':all_steps[9:13],'loopD':all_steps[13:]}
    for i in all_steps[0:2]:
        key = i
        dict_steps={key:'loopA'}
        dict_loops.update(dict_steps)
    for i in all_steps[2:9]:
        key = i
        dict_steps={key:'loopB'}
        dict_loops.update(dict_steps)
    for i in all_steps[9:13]:
        key = i
        dict_steps={key:'loopC'}
        dict_loops.update(dict_steps)
    for i in all_steps[13:]:
        key = i
        dict_steps={key:'loopD'}
        dict_loops.update(dict_steps)
    return dict_loops
    
def feasible_dataset():
    time_max = xlrd.open_workbook('time_range.xlsx').sheet_by_name('Sheet1').col_values(2)[1:]
    return time_max

def comparison():
    dict_ref = statistical_values()
    dict_data = not_feasible_step()
    dict_loops = loop_and_steps()
    
    ref = list(dict_ref.values())
    data = list(dict_data.values())
    time_max = feasible_dataset()
    steps = ['type N','type A','type B','type C','type D','type E','type F','type G','type H']
    book = xlwt.Workbook() 
    sheet1 = book.add_sheet('step_types') 
    for i in range(0,len(steps)):
        sheet1.write(0,i,steps[i])
        book.save('step_types.xls')
    m = 0
    n = 0
    c = 1
    for i in range(0,len(data)):
        val = data[i]
        count = 0
        count = val.count(' ')
        if count <= 0.2*len(val):
            name_N = list(dict_data.keys())[i]
            sheet1.write(c,0,name_N)
            book.save('step_types.xls')
            c=c+1
            print("Step [%str] is a type N step" %(name_N))
        else:
            for j in val:
                if j != ' ':
                    num=int(j)
                    diff = (num-time_max[n])/time_max[n]
                    if diff < 0.5:
                        m=m+1
            #if m > 0.7*len(val):
        n = n+1

    
if __name__ =="__main__":
    loop_and_steps()
