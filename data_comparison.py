#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Sep 11 10:44:44 2020

@author: ziling
"""

import step_evaluation1
import simpy
from simpy.events import AllOf
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
        for j in range(0,len(step_ana.row_values(0)[1:-3])):
            value.append(step_ana.row_values(n)[j+1])
            #m = m+1
            for s in value:
                if s != ' ':
                    value = [float(s) if x == s else x for x in value]
                dict_anadata = {i:value}
        dict_data.update(dict_anadata)
        n = n+1
    return dict_data

def statistical_values():
    #get the analyzable steps
    step_ana = xlrd.open_workbook(step_evaluation1.excel_table_byname(file= 'Data_set.xlsx', colnameindex=0,         by_name=u'Sheet1')).sheet_by_name('new_dataset')
    key = step_ana.row_values(0)[-3:]
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

def comparison():
    dict_ref = statistical_values()
    dict_data = not_feasible_step()
    ref = list(dict_ref.values())
    data = list(dict_data.values())
    # print(data)
    for i in range(0,len(data)):
        val = data[i]
        print(val)
        #for j in val:
            
    
if __name__ =="__main__":
    comparison()