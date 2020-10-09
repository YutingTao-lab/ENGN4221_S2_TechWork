#!/usr/bin/env python3
# -*- coding: utf-8 -*-


import xlrd
import xlwt
import numpy as np
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
    data = open_excel(file) #open excel file
    table = data.sheet_by_name(by_name) #obtain data from excel file
    book = xlwt.Workbook() #create a new excel file
    sheet1 = book.add_sheet('sheet1')
    col0 = ['Stage Loop','Loop A', 'Loop B','Loop C', 'Loop D']
    i = 0
    for i in range(0,len(col0)):
        sheet1.write(i,0,str(col0[i]))
        book.save('looparrange.xls')
    row0=table.row_values(0)
    row0.pop(1)
    j = 0
    for j in range(1,len(row0)-3):
        sheet1.write(0,j,str(row0[j]))
        book.save('looparrange.xls')
    sheet1.write(0,len(row0)-3,'frequency')
    book.save('looparrange.xls')
#obtain the matrix of data with different set
    set_matrix=[]
    for row in range (1,table.nrows):
        _row = []
        for col in range (1,table.ncols-4):
            _row.append(table.cell_value(row,col+1))
        set_matrix.append(_row)
    set_matrix_array = np.array(set_matrix)
    A = np.sum(set_matrix_array[0:2], axis = 0)
    B = np.sum(set_matrix_array[2:9], axis = 0)
    C = np.sum(set_matrix_array[9:13], axis = 0)
    D = np.sum(set_matrix_array[13:], axis = 0)
    loop_matrix = np.vstack((A,B,C,D))
#obtain frequency
    trans_loop = loop_matrix.transpose()
    loc = []
    for n in range(0,len(trans_loop)):
        a = trans_loop[n]
        b = np.where(a == a.max())
        c = b[0]
        loc.append(c[0])
    AF = 0
    BF = 0
    CF = 0
    DF = 0
    for m in range(0,len(loc)):
        if loc[n] == 0:
            AF = AF + 1
        if loc[n] == 1:
            BF = BF + 1
        if loc[n] == 2:
            CF = CF + 1
        if loc[n] == 3:
            DF == DF + 1
    frequency = np.array([AF,BF,CF,DF])

    trans_loop_matrix_with_freq = np.vstack((trans_loop,frequency))
    loop_matrix_with_freq = trans_loop_matrix_with_freq.transpose()

    [h, l] = loop_matrix_with_freq.shape
    for i in range(h):
        for j in range(l):
            sheet1.write(i+1, j+1, loop_matrix_with_freq[i, j])
            book.save('looparrange.xls')
    
    max_freq = np.where(frequency == frequency.max())
    if np.any(max_freq[0]==[0]):
        # print('loop A has max frequency')
        return 'A'
    if np.any(max_freq[0]==[1]):
        # print('loop B has max frequency')
        return 'B'
    if np.any(max_freq[0]==[2]):
        # print('loop C has max frequency')
        return 'C'
    if np.any(max_freq[0]==[3]):
        # print('loop D has max frequency')
        return 'D'
        
    

    

# if __name__ =="__main__":
#   excel_table_byname(file= 'Data_estimation.xls', colnameindex=0, by_name=u'sheet1')
        
    
