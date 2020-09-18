#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Sep 18 20:06:26 2020

@author: ziling
"""

import step_evaluation1
import four_loops_arrange
import sympy
import Data_Estimation
# from sympy.events import AllOf
import pandas as pd
import  xdrlib ,sys
import xlrd
import xlwt
import math

def data_step():
    #get the analyzable steps
    step_ana1 = step_evaluation1.excel_table(file= 'Data_estimation.xls', colnameindex=0, by_name=u'sheet1')
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

def define_error_data():
    r=5
    pmin=0.00
    pmax=0.00
    # i=0
    s3=[1.15,1.46,1.67,1.82,1.94,2.03,2.11,2.18,2.23,2.29,2.33,2.37,2.41,2.44,2.47,2.50,2.53,2.56,2.58,2.60,2.62,2.64,2.66,2.68,2.7,2.71,2.73,2.75]
    dict_data=data_step()
    steps=list(dict_data.keys())
    steps_values = list(dict_data.values())
    for j in range(0,len(steps)):
        p1=[]
        error=[]
        samle=0
        n=30
        num = steps_values[j]
        num_set = num[0:-7]
        # x = num[-6] #mean value
        ss = num[-4] #std value
        for i in range(0,n):
            p1.insert(i,num[i])
            # print(i,p,num)
        p=p1
        p.sort(reverse=False)
        if len(p)/2 is int:
            x = (p[len(p)/2]+p[(len(p)/2)+1])/2
        else:
            x = int(len(p)/2)+1
            
        for a in p:
            Gi = (a-x)/ss
            if Gi > s3[n-3]:
                error.insert(-1,a)
        print("异常数据是：",error,len(error))

    # return error
    
if __name__ =="__main__":
    define_error_data()