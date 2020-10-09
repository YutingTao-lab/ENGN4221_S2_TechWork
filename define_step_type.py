#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Sep 27 14:37:10 2020

@author: ziling
"""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Sep  9 17:29:46 2020

@author: ziling
"""


import four_loops_arrange
import xlrd
from xlutils.copy import copy


def data_step():
    #get the analyzable steps
    # step_ana1 = step_evaluation1.excel_table(file= 'Data_estimation.xls', colnameindex=0, by_name=u'sheet1')
    newdata = xlrd.open_workbook('new_dataset.xls')
    step_ana = newdata.sheet_by_name('new_dataset')
    key = step_ana.col_values(0)[1:]
    n = 1
    dict_data = {}
    for i in key:
        #m = 1
        value =[]
        # print(step_ana.row_values(1))
        for j in range(0,len(step_ana.row_values(0)[1:])):
            m = float(step_ana.row_values(n)[j+1])
            # print(m)
            value.insert(j,m)
            # print(m,value)
            #m = m+1
            dict_anadata = {i:value}
        dict_data.update(dict_anadata)
        #print(dict_anadata)
        n = n+1
    return dict_data

def low_var_data():
    time_lmax = xlrd.open_workbook('ideal_range.xls').sheet_by_name('sheet1').col_values(1)[1:]
    return time_lmax

def High_var_data():
    time_hmax = xlrd.open_workbook('ideal_range.xls').sheet_by_name('sheet1').col_values(2)[1:]
    return time_hmax

def loop_and_steps():
    all_steps = xlrd.open_workbook('Data_estimation.xls').sheet_by_name('sheet1').col_values(0)
    dict_loops={}
    # dict_loops = {'loopA':all_steps[0:2],'loopB':all_steps[2:9],'loopC':all_steps[9:13],'loopD':all_steps[13:]}
    for i in all_steps[0:2]:
        key = i
        dict_steps={key:'A'}
        dict_loops.update(dict_steps)
    for i in all_steps[2:9]:
        key = i
        dict_steps={key:'B'}
        dict_loops.update(dict_steps)
    for i in all_steps[9:13]:
        key = i
        dict_steps={key:'C'}
        dict_loops.update(dict_steps)
    for i in all_steps[13:]:
        key = i
        dict_steps={key:'D'}
        dict_loops.update(dict_steps)
    return dict_loops

def define_error_data():    
    # get_ideal_range()
    time_lmax = low_var_data()
    time_hmax = High_var_data()
    dict_data=data_step()
    dict_loops = loop_and_steps()
    steps=list(dict_data.keys())
    steps_values = list(dict_data.values())
    dict_loops=loop_and_steps()
    maxloop = four_loops_arrange.excel_table_byname(file= 'Data_estimation.xls', colnameindex=0, by_name=u'sheet1')
    time_set1 = xlrd.open_workbook('ideal_range.xls')
    
    #create new sheet to describe the type of steps
    time_range1 = time_set1.sheet_by_name('sheet1')
    time_set = copy(time_set1)
    step_type = time_set.get_sheet(1)
    step_type.write(0,0,'improvable steps')
    step_type.write(0,1,'step types')
    time_set.save('ideal_range.xls')
    
    for i in range(0,len(steps)):
        step_type.write(i+1,0,steps[i])
        time_set.save('ideal_range.xls')
    #Define error data and problem types
    #Evaluate and classify steps that can be improved
    r = 0
    for j in range(0,len(steps)):
        r = r+1
        j_steps = steps[j]
        lmax = time_lmax[j]
        hmax = time_hmax[j]
        count1=0
        count2=0
        for i in steps_values[j]:
            if i >= lmax:
                if i < hmax:
                    count1 = count1+1
                else:
                    count2 = count2+1
        if count1+count2 <= 0.1*len(steps_values[j]):
            step_type.write(r,1,'type N')
            time_set.save('ideal_range.xls')
            # print('Step',j_steps,'is type N')
        else:
            if count2 != 0:
                if count2>=count1:
                    if dict_loops[j_steps] == maxloop:
                        step_type.write(r,1,'type E')
                        time_set.save('ideal_range.xls')
                        # print('Step',j_steps,'is type E')
                    else:
                        step_type.write(r,1,'type F')
                        time_set.save('ideal_range.xls')
                        # print('Step',j_steps,'is type F')
                else:
                    if dict_loops[j_steps] == maxloop:
                        step_type.write(r,1,'type G')
                        time_set.save('ideal_range.xls')
                        # print('Step',j_steps,'is type G')
                    else:
                        step_type.write(r,1,'type H')
                        time_set.save('ideal_range.xls')
                        # print('Step',j_steps,'is type H')
            else:
                if count1>=0.5*len(steps_values[j]):
                    if dict_loops[j_steps] == maxloop:
                        step_type.write(r,1,'type A')
                        time_set.save('ideal_range.xls')
                        # print('Step',j_steps,'is type A')
                    else:
                        step_type.write(r,1,'type B')
                        time_set.save('ideal_range.xls')
                        # print('Step',j_steps,'is type B')
                else:
                    if dict_loops[j_steps] == maxloop:
                        step_type.write(r,1,'type C')
                        time_set.save('ideal_range.xls')
                        # print('Step',j_steps,'is type C')
                    else:
                        step_type.write(r,1,'type D')
                        time_set.save('ideal_range.xls')
                        # print('Step',j_steps,'is type D')
    time_set = xlrd.open_workbook('ideal_range.xls')
    step_type=time_set.sheet_by_name('step_type')
    keys = step_type.col_values(0)[1:]
    values = step_type.col_values(-1)[1:]
    dict_step = {}
    for i in range(len(keys)):
        step = keys[i]
        step_type = values[i]
        dict_step1 = {step:step_type}
        dict_step.update(dict_step1)
    return dict_step
