#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Sep  9 17:29:46 2020

@author: ziling
"""


import step_evaluation1
import Data_Estimation
import time_range
import define_step_type
import simulation_full_stage
import xlrd
from xlutils.copy import copy

Data_Estimation.excel_table_byname()
time_range.excel_table_byname(file= 'Data_estimation.xls', colnameindex=0, by_name=u'sheet1')
step_ana1 = step_evaluation1.excel_table(file= 'Data_estimation.xls', colnameindex=0, by_name=u'sheet1')
newdata = xlrd.open_workbook('new_dataset.xls')
step_ana = newdata.sheet_by_name('new_dataset')
steps = step_ana.col_values(0)[1:]
dict_step = {}
step_value = []
time_set1 = xlrd.open_workbook('ideal_range.xls')
time_set = copy(time_set1)
#create new sheet to describe the type of steps

step_type = time_set.add_sheet('step_type')
step_type = time_set.get_sheet(1)
step_type.write(0,0,'improvable steps')
step_type.write(0,1,'step types')
time_set.save('ideal_range.xls')

for i in range(0,20):
    values = list(define_step_type.define_error_data().values())
    for j in range(len(steps)):
        # print(values,j)
        step_value = values[j]
        step = steps[j]
        if step in dict_step:
            dict_step.get(step).append(step_value)
        else:
            dict_step.setdefault(step,[]).append(step_value)
k=1
for i in steps:
    N = 0
    A= 0
    B = 0
    C = 0
    D = 0
    E = 0
    F = 0
    G = 0
    H = 0
    get_value = dict_step[i]
    # print(get_value)
    for j in range(len(get_value)):
        if get_value[j] == 'type A':
            A = A +1
        if get_value[j] == 'type B':
            B = B +1
        if get_value[j] == 'type C':
            C = C +1
        if get_value[j] == 'type D':
            D = D +1
        if get_value[j] == 'type E':
            E = E +1
        if get_value[j] == 'type F':
            F = F +1
        if get_value[j] == 'type G':
            G = G +1
        if get_value[j] == 'type H':
            H = H +1
        if get_value[j] == 'type N':
            N = N +1
    #print(A)
    type_list = ['type A','type B','type C','type D','type E','type F','type G','type H','type N']
    frequency = [A,B,C,E,F,G,H,N]
    max_position=[frequency.index(max(frequency))]
    m=1
    step_type.write(k,0,i)
    if len(max_position) == 1:
        step_t = type_list[max_position[0]]
        step_type.write(k,1,step_t)
        time_set.save('ideal_range.xls')
        print(i,'is',step_t)
    else:
        for n in range(len(max_position)):
            step_t = type_list[max_position[n]]
            step_type.write(k,m,step_t)
            time_set.save('ideal_range.xls')
            m=m+1
        print(i,'is',step_t)
    k=k+1

simulation_full_stage.total_time_after_reduce()



