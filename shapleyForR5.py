#!/usr/bin/env python
from itertools import combinations
import xlrd
import xlwt
import re
import random
import numpy as np
import pandas as pd
from sklearn.metrics import mean_squared_error
a=["M","D","R","C","H"];
# b=[0.74421029, 0.30133586, 0.72358195,  0.547016, 0.43886731] #test
b=[0.81340156, 0.62245651, 0.05931,  0.46608105, 0.98067221]#test3
# b=[0.84502405, 0.36937625, 0.02653019,  0.5268023, 0.64680392]#test2
# b=[0.95487932,0.04479157,0.80758104,0.68115175,0.39437373]#test4
# b=[0.61259584,0.88131645,0.08123287,0.38380541,0.1926135]#test1


def randoms(x):

    test=["M", "R", "D", "C","H", "MM", "MD", "MC","MH" "DD", "DC","DH", "CC",
          "CH","HH","MR", "DR", "RR", "RC","RH",
     "MMM", "MMD", "MMR", "MMC", "MMH","MDD", "MDR", "MDC","MDH", "MRR", "MRC", "MRH","MCC",
     "MCH","MHH","DDD", "DDR", "DDC", "DDH","DRR", "DRC","DRH", "DCC","DCH","DHH",
          "RRC", "RRH","RRR", "RCC","RCH","RHH", "CCC","CCH","CHH","HHH",
          "MMMM", "MMMD", "MMMR", "MMMC", "MMMH", "MMDD", "MMDR", "MMDC", "MMDH", "MMRR", "MMRC", "MMRH", "MMCC",
          "MMCH", "MMHH", "MDDD", "MDDR", "MDDC", "MDDH", "MDRR", "MDRC", "MDRH", "MDCC", "MDCH", "MDHH",
          "MRRC", "MRRH", "MRRR", "MRCC", "MRCH", "MRHH", "MCCC", "MCCH", "MCHH", "MHHH",
          "DDDD","DDDR","DDDC","DDDH","DDRR","DDRC","DDRH","DDCC","DDCH","DDHH","DRRR","DRRC","DRRH",
          "DRCC","DRCH","DRHH","DCCC","DCCH","DCHH","DHHH","RRRR","RRRC","RRRH","RRCC","RRCH","RRHH",
          "RCCC","RCCH","RCHH","RHHH","CCCC","CCCH","CCHH","CHHH","HHHH"
          ]
    F=random.sample(test, x)
    P = np.random.uniform(low=0, high=1, size=x)
    return F,P
def openexcel(workbook,lab):
    wb1 = workbook.add_sheet(lab, cell_overwrite_ok=True)

    return wb1

def findvalue(sub,test):
    if sub=='':
        return 0
    table = xlrd.open_workbook("16or2.xlsx")
    sheet = table.sheet_by_name("test"+test.__str__())
    row_count = sheet.nrows
    column_count = sheet.ncols
    for j in range(0, column_count):
        for i in range(0, row_count):
            cell = sheet.cell(i, j)
            cell = str(cell).split(":")
            want = ''.join(re.findall(r'[A-Za-z]', str(cell[1])))
            if want==sub:
                yy = sheet.cell(i, 1).value
                return yy
def findvalueD(sub,test):

    table = xlrd.open_workbook("test2.xlsx")
    sheet = table.sheet_by_name("Sheet2")
    row_count = sheet.nrows
    column_count = sheet.ncols
    for j in range(0, column_count):
        for i in range(0, row_count):
            cell = sheet.cell(i, j)
            cell = str(cell).split(":")
            want = ''.join(re.findall(r'[A-Za-z]', str(cell[1])))
            if want==sub:
                yy = sheet.cell(i, test).value
                return yy
def findvalueR(sub,test):

    table = xlrd.open_workbook("test2.xlsx")
    sheet = table.sheet_by_name("Sheet3")
    row_count = sheet.nrows
    column_count = sheet.ncols
    for j in range(0, column_count):
        for i in range(0, row_count):
            cell = sheet.cell(i, j)
            cell = str(cell).split(":")
            want = ''.join(re.findall(r'[A-Za-z]', str(cell[1])))
            if want==sub:
                yy = sheet.cell(i, test).value
                return yy


def findvalueC(sub,test):
    table = xlrd.open_workbook("test2.xlsx")
    sheet = table.sheet_by_name("Sheet4")
    row_count = sheet.nrows
    column_count = sheet.ncols
    for j in range(0, column_count):
        for i in range(0, row_count):
            cell = sheet.cell(i, j)
            cell = str(cell).split(":")
            want = ''.join(re.findall(r'[A-Za-z]', str(cell[1])))
            if want == sub:
                yy = sheet.cell(i, test).value
                return yy
def findvalueH(sub,test):
    table = xlrd.open_workbook("test2.xlsx")
    sheet = table.sheet_by_name("Sheet5")
    row_count = sheet.nrows
    column_count = sheet.ncols
    for j in range(0, column_count):
        for i in range(0, row_count):
            cell = sheet.cell(i, j)
            cell = str(cell).split(":")
            want = ''.join(re.findall(r'[A-Za-z]', str(cell[1])))
            if want == sub:
                yy = sheet.cell(i, test).value
                return yy

def sampleM(r11, mm1, mm2, mm3, mm4, sum1M, RM, CM,FM):
    r0=[]
    r1=[]
    r2 = []
    r3 = []
    r4=[]
    tt=0
    kk = 0
    j = 0
    uu = 0
    for i in range(0,len(r11)):
        if len(r11[i])==0:
            aa = b[0]
            r0.append(aa)

        if len(r11[i])==1:
            bb = find(r11[i], mm1, sum1M)
            r1.append(bb)
            tt=tt+1
        if len(r11[i])==2:
            cc = find(r11[i], mm2, RM)
            r2.append(cc)
            kk=kk+1
        if len(r11[i]) == 3:
            dd = find(r11[i], mm3, CM)
            r3.append(dd)
            j = j + 1
        if len(r11[i]) == 4:
            ee = find(r11[i], mm4, FM)
            r4.append(ee)
            uu = uu + 1
    if len(r1)==0 and len(r2) == 0 and len(r3) == 0 and len(r4) == 0 and len(r0) == 0:
        gg=0
    elif len(r1)==0 and len(r2) == 0 and len(r3) == 0 and len(r4) == 0:
        gg = (b[0]) / 5
    elif len(r0) == 0 and len(r2) == 0 and len(r3) == 0 and len(r4) == 0:
        gg = (sum(r1) / tt) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r3) == 0 and len(r4) == 0:
        gg = (sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) == 0 and len(r4) == 0:
        gg = (sum(r3) / j) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu) / 5

    elif len(r2) == 0 and len(r3) == 0 and len(r0) != 0 and len(r1) != 0 and len(r4) == 0:
        gg = (sum(r1) / tt + b[0]) / 5
    elif len(r0) == 0 and len(r1) != 0 and len(r2) != 0 and len(r3) == 0 and len(r4) == 0:
        gg = (sum(r1) / tt + sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r1) != 0 and len(r2) == 0 and len(r3) != 0 and len(r4) == 0:
        gg = (sum(r1) / tt + sum(r3) / j) / 5
    elif len(r0) != 0 and len(r1) == 0 and len(r2) != 0 and len(r3) == 0 and len(r4) == 0:
        gg = (b[0] + sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) != 0 and len(r3) != 0 and len(r4) == 0:
        gg = (sum(r3) / j + sum(r2) / kk) / 5
    elif len(r0) != 0 and len(r1) == 0 and len(r2) == 0 and len(r3) != 0 and len(r4) == 0:
        gg = (sum(r3) / j + b[0]) / 5

    elif len(r0) != 0 and len(r1) == 0 and len(r2) ==  0 and len(r3) ==  0 and len(r4) != 0:
        gg = (sum(r4) / uu + b[0]) / 5
    elif len(r0) == 0 and len(r1) != 0 and len(r2) ==  0 and len(r3) ==  0 and len(r4) != 0:
        gg = (sum(r4) / uu + sum(r1) / tt) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) !=  0 and len(r3) ==  0 and len(r4) != 0:
        gg = (sum(r2) / kk + sum(r4) / uu) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) ==  0 and len(r3) !=  0 and len(r4) != 0:
        gg = (sum(r3) / j + sum(r4) / uu) /5

    elif len(r2) == 0 and len(r4) == 0 :
        gg = (sum(r1) / tt + b[0] + sum(r3) / j) / 5

    elif len(r3) == 0 and len(r4) == 0:
        gg = (sum(r1) / tt + b[0] + sum(r2) / kk) / 5

    elif len(r1) == 0 and len(r4) == 0:
        gg = (sum(r3) / j + b[0] + sum(r2) / kk) / 5

    elif len(r0) == 0 and len(r4) == 0:
        gg = (sum(r3) / j + sum(r1) / tt  + sum(r2) / kk) / 5
    elif len(r2) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt  + b[0]) / 5
    elif len(r1) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk  + b[0]) / 5
    elif len(r1) == 0 and len(r2) == 0:
        gg = (sum(r4) / uu + sum(r3) / j  + b[0]) / 5
    elif len(r0) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt + sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r2) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt + sum(r3) / j) /5
    elif len(r0) == 0 and len(r1) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + sum(r3) / j) /5
    elif len(r0) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + sum(r3) / j+ sum(r1) / tt) /5
    elif len(r1) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + sum(r3) / j+ b[0]) /5
    elif len(r2) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt + sum(r3) / j + b[0]) / 5
    elif len(r3) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + b[0]+ sum(r1) / tt) /5
    elif len(r4) == 0:
        gg = (sum(r3) / j + sum(r2) / kk + b[0]+ sum(r1) / tt) /5
    else:

        gg = (sum(r1) / tt + b[0] + sum(r2) / kk + sum(r3) / j+sum(r4) / uu ) / 5

    return gg
def sampleD(r11, mm1, mm2, mm3, mm4, sum1M, RM, CM,FM):
    r0=[]
    r1=[]
    r2 = []
    r3 = []
    r4=[]
    tt=0
    kk = 0
    j = 0
    uu = 0
    for i in range(0,len(r11)):
        if len(r11[i])==0:
            aa = b[1]
            r0.append(aa)

        if len(r11[i])==1:
            bb = find(r11[i], mm1, sum1M)
            r1.append(bb)
            tt=tt+1
        if len(r11[i])==2:
            cc = find(r11[i], mm2, RM)
            r2.append(cc)
            kk=kk+1
        if len(r11[i]) == 3:
            dd = find(r11[i], mm3, CM)
            r3.append(dd)
            j = j + 1
        if len(r11[i]) == 4:
            ee = find(r11[i], mm4, FM)
            r4.append(ee)
            uu = uu + 1
    if len(r1) == 0 and len(r2) == 0 and len(r3) == 0 and len(r4) == 0 and len(r0) == 0:
        gg = 0
    elif len(r1)==0 and len(r2) == 0 and len(r3) == 0 and len(r4) == 0:
        gg = (b[1]) / 5
    elif len(r0) == 0 and len(r2) == 0 and len(r3) == 0 and len(r4) == 0:
        gg = (sum(r1) / tt) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r3) == 0 and len(r4) == 0:
        gg = (sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) == 0 and len(r4) == 0:
        gg = (sum(r3) / j) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu) / 5

    elif len(r2) == 0 and len(r3) == 0 and len(r0) != 0 and len(r1) != 0 and len(r4) == 0:
        gg = (sum(r1) / tt + b[1]) / 5
    elif len(r0) == 0 and len(r1) != 0 and len(r2) != 0 and len(r3) == 0 and len(r4) == 0:
        gg = (sum(r1) / tt + sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r1) != 0 and len(r2) == 0 and len(r3) != 0 and len(r4) == 0:
        gg = (sum(r1) / tt + sum(r3) / j) / 5
    elif len(r0) != 0 and len(r1) == 0 and len(r2) != 0 and len(r3) == 0 and len(r4) == 0:
        gg = (b[1] + sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) != 0 and len(r3) != 0 and len(r4) == 0:
        gg = (sum(r3) / j + sum(r2) / kk) /5
    elif len(r0) != 0 and len(r1) == 0 and len(r2) == 0 and len(r3) != 0 and len(r4) == 0:
        gg = (sum(r3) / j + b[1]) / 5

    elif len(r0) != 0 and len(r1) == 0 and len(r2) ==  0 and len(r3) ==  0 and len(r4) != 0:
        gg = (sum(r4) / uu + b[1]) / 5
    elif len(r0) == 0 and len(r1) != 0 and len(r2) ==  0 and len(r3) ==  0 and len(r4) != 0:
        gg = (sum(r4) / uu + sum(r1) / tt) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) !=  0 and len(r3) ==  0 and len(r4) != 0:
        gg = (sum(r2) / kk + sum(r4) / uu) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) ==  0 and len(r3) !=  0 and len(r4) != 0:
        gg = (sum(r3) / j + sum(r4) / uu) / 5

    elif len(r2) == 0 and len(r4) == 0 :
        gg = (sum(r1) / tt + b[1] + sum(r3) / j) / 5

    elif len(r3) == 0 and len(r4) == 0:
        gg = (sum(r1) / tt + b[1] + sum(r2) / kk) / 5

    elif len(r1) == 0 and len(r4) == 0:
        gg = (sum(r3) / j + b[1] + sum(r2) / kk) / 5

    elif len(r0) == 0 and len(r4) == 0:
        gg = (sum(r3) / j + sum(r1) / tt  + sum(r2) / kk) / 5
    elif len(r2) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt  + b[1]) / 5
    elif len(r1) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk  + b[1]) / 5
    elif len(r1) == 0 and len(r2) == 0:
        gg = (sum(r4) / uu + sum(r3) / j  + b[1]) / 5
    elif len(r0) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt + sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r2) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt + sum(r3) / j) / 5
    elif len(r0) == 0 and len(r1) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + sum(r3) / j) / 5
    elif len(r0) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + sum(r3) / j + sum(r1) / tt) / 5
    elif len(r1) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + sum(r3) / j + b[1]) / 5
    elif len(r2) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt + sum(r3) / j + b[1]) / 5
    elif len(r3) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + b[1] + sum(r1) / tt) / 5
    elif len(r4) == 0:
        gg = (sum(r3) / j + sum(r2) / kk + b[1] + sum(r1) / tt) / 5
    else:
        gg = (sum(r1) / tt + b[1] + sum(r2) / kk + sum(r3) / j+sum(r4) / uu) / 5

    return gg
def sampleR(r11, mm1, mm2, mm3, mm4, sum1M, RM, CM,FM):
    r0=[]
    r1=[]
    r2 = []
    r3 = []
    r4=[]
    tt=0
    kk = 0
    j = 0
    uu = 0
    for i in range(0,len(r11)):
        if len(r11[i])==0:
            aa = b[2]
            r0.append(aa)

        if len(r11[i])==1:
            bb = find(r11[i], mm1, sum1M)
            r1.append(bb)
            tt=tt+1
        if len(r11[i])==2:
            cc = find(r11[i], mm2, RM)
            r2.append(cc)
            kk=kk+1
        if len(r11[i]) == 3:
            dd = find(r11[i], mm3, CM)
            r3.append(dd)
            j = j + 1
        if len(r11[i]) == 4:
            ee = find(r11[i], mm4, FM)
            r4.append(ee)
            uu = uu + 1
    if len(r1) == 0 and len(r2) == 0 and len(r3) == 0 and len(r4) == 0 and len(r0) == 0:
        gg = 0
    elif len(r1)==0 and len(r2) == 0 and len(r3) == 0 and len(r4) == 0:
        gg = (b[2]) / 5
    elif len(r0) == 0 and len(r2) == 0 and len(r3) == 0 and len(r4) == 0:
        gg = (sum(r1) / tt) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r3) == 0 and len(r4) == 0:
        gg = (sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) == 0 and len(r4) == 0:
        gg = (sum(r3) / j) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu) / 5

    elif len(r2) == 0 and len(r3) == 0 and len(r0) != 0 and len(r1) != 0 and len(r4) == 0:
        gg = (sum(r1) / tt + b[2]) / 5
    elif len(r0) == 0 and len(r1) != 0 and len(r2) != 0 and len(r3) == 0 and len(r4) == 0:
        gg = (sum(r1) / tt + sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r1) != 0 and len(r2) == 0 and len(r3) != 0 and len(r4) == 0:
        gg = (sum(r1) / tt + sum(r3) / j) / 5
    elif len(r0) != 0 and len(r1) == 0 and len(r2) != 0 and len(r3) == 0 and len(r4) == 0:
        gg = (b[2] + sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) != 0 and len(r3) != 0 and len(r4) == 0:
        gg = (sum(r3) / j + sum(r2) / kk) / 5
    elif len(r0) != 0 and len(r1) == 0 and len(r2) == 0 and len(r3) != 0 and len(r4) == 0:
        gg = (sum(r3) / j + b[2]) / 5

    elif len(r0) != 0 and len(r1) == 0 and len(r2) ==  0 and len(r3) ==  0 and len(r4) != 0:
        gg = (sum(r4) / uu + b[2]) / 5
    elif len(r0) == 0 and len(r1) != 0 and len(r2) ==  0 and len(r3) ==  0 and len(r4) != 0:
        gg = (sum(r4) / uu + sum(r1) / tt) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) !=  0 and len(r3) ==  0 and len(r4) != 0:
        gg = (sum(r2) / kk + sum(r4) / uu) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) ==  0 and len(r3) !=  0 and len(r4) != 0:
        gg = (sum(r3) / j + sum(r4) / uu) / 5

    elif len(r2) == 0 and len(r4) == 0 :
        gg = (sum(r1) / tt + b[2] + sum(r3) / j) / 5

    elif len(r3) == 0 and len(r4) == 0:
        gg = (sum(r1) / tt + b[2] + sum(r2) / kk) / 5

    elif len(r1) == 0 and len(r4) == 0:
        gg = (sum(r3) / j + b[2] + sum(r2) / kk) / 5

    elif len(r0) == 0 and len(r4) == 0:
        gg = (sum(r3) / j + sum(r1) / tt  + sum(r2) / kk) / 5
    elif len(r2) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt  + b[2]) / 5
    elif len(r1) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk  + b[2]) / 5
    elif len(r1) == 0 and len(r2) == 0:
        gg = (sum(r4) / uu + sum(r3) / j  + b[2]) / 5
    elif len(r0) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt + sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r2) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt + sum(r3) / j) / 5
    elif len(r0) == 0 and len(r1) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + sum(r3) / j) / 5
    elif len(r0) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + sum(r3) / j + sum(r1) / tt) / 5
    elif len(r1) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + sum(r3) / j + b[2]) / 5
    elif len(r2) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt + sum(r3) / j + b[2]) / 5
    elif len(r3) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + b[2] + sum(r1) / tt) / 5
    elif len(r4) == 0:
        gg = (sum(r3) / j + sum(r2) / kk + b[2] + sum(r1) / tt) / 5
    else:
        gg = (sum(r1) / tt + b[2] + sum(r2) / kk + sum(r3) / j+sum(r4) / uu) /5

    return gg
def sampleC(r11, mm1, mm2, mm3, mm4, sum1M, RM, CM,FM):
    r0=[]
    r1=[]
    r2 = []
    r3 = []
    r4=[]
    tt=0
    kk = 0
    j = 0
    uu = 0
    for i in range(0,len(r11)):
        if len(r11[i])==0:
            aa = b[3]
            r0.append(aa)

        if len(r11[i])==1:
            bb = find(r11[i], mm1, sum1M)
            r1.append(bb)
            tt=tt+1
        if len(r11[i])==2:
            cc = find(r11[i], mm2, RM)
            r2.append(cc)
            kk=kk+1
        if len(r11[i]) == 3:
            dd = find(r11[i], mm3, CM)
            r3.append(dd)
            j = j + 1
        if len(r11[i]) == 4:
            ee = find(r11[i], mm4, FM)
            r4.append(ee)
            uu = uu + 1
    if len(r1) == 0 and len(r2) == 0 and len(r3) == 0 and len(r4) == 0 and len(r0) == 0:
        gg = 0
    elif len(r1)==0 and len(r2) == 0 and len(r3) == 0 and len(r4) == 0:
        gg = (b[3]) / 5
    elif len(r0) == 0 and len(r2) == 0 and len(r3) == 0 and len(r4) == 0:
        gg = (sum(r1) / tt) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r3) == 0 and len(r4) == 0:
        gg = (sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) == 0 and len(r4) == 0:
        gg = (sum(r3) / j) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu) / 5

    elif len(r2) == 0 and len(r3) == 0 and len(r0) != 0 and len(r1) != 0 and len(r4) == 0:
        gg = (sum(r1) / tt + b[3]) / 5
    elif len(r0) == 0 and len(r1) != 0 and len(r2) != 0 and len(r3) == 0 and len(r4) == 0:
        gg = (sum(r1) / tt + sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r1) != 0 and len(r2) == 0 and len(r3) != 0 and len(r4) == 0:
        gg = (sum(r1) / tt + sum(r3) / j) / 5
    elif len(r0) != 0 and len(r1) == 0 and len(r2) != 0 and len(r3) == 0 and len(r4) == 0:
        gg = (b[3] + sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) != 0 and len(r3) != 0 and len(r4) == 0:
        gg = (sum(r3) / j + sum(r2) / kk) / 5
    elif len(r0) != 0 and len(r1) == 0 and len(r2) == 0 and len(r3) != 0 and len(r4) == 0:
        gg = (sum(r3) / j + b[3]) / 5

    elif len(r0) != 0 and len(r1) == 0 and len(r2) ==  0 and len(r3) ==  0 and len(r4) != 0:
        gg = (sum(r4) / uu + b[3]) / 5
    elif len(r0) == 0 and len(r1) != 0 and len(r2) ==  0 and len(r3) ==  0 and len(r4) != 0:
        gg = (sum(r4) / uu + sum(r1) / tt) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) !=  0 and len(r3) ==  0 and len(r4) != 0:
        gg = (sum(r2) / kk + sum(r4) / uu) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) ==  0 and len(r3) !=  0 and len(r4) != 0:
        gg = (sum(r3) / j + sum(r4) / uu) / 5

    elif len(r2) == 0 and len(r4) == 0 :
        gg = (sum(r1) / tt + b[3] + sum(r3) / j) / 5

    elif len(r3) == 0 and len(r4) == 0:
        gg = (sum(r1) / tt + b[3] + sum(r2) / kk) / 5

    elif len(r1) == 0 and len(r4) == 0:
        gg = (sum(r3) / j + b[3] + sum(r2) / kk) / 5

    elif len(r0) == 0 and len(r4) == 0:
        gg = (sum(r3) / j + sum(r1) / tt  + sum(r2) / kk) / 5
    elif len(r2) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt  + b[3]) / 5
    elif len(r1) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk  + b[3]) / 5
    elif len(r1) == 0 and len(r2) == 0:
        gg = (sum(r4) / uu + sum(r3) / j  + b[3]) / 5
    elif len(r0) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt + sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r2) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt + sum(r3) / j) / 5
    elif len(r0) == 0 and len(r1) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + sum(r3) / j) / 5
    elif len(r0) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + sum(r3) / j + sum(r1) / tt) / 5
    elif len(r1) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + sum(r3) / j + b[3]) / 5
    elif len(r2) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt + sum(r3) / j + b[3]) / 5
    elif len(r3) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + b[3] + sum(r1) / tt) / 5
    elif len(r4) == 0:
        gg = (sum(r3) / j + sum(r2) / kk + b[3] + sum(r1) / tt) / 5
    elif len(r0) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + sum(r3) / j + sum(r1) / tt) / 5
    elif len(r1) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + sum(r3) / j + b[3]) / 5
    elif len(r2) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt + sum(r3) / j + b[3]) / 5
    elif len(r3) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + b[3] + sum(r1) / tt) / 5
    elif len(r4) == 0:
        gg = (sum(r3) / j + sum(r2) / kk + b[3] + sum(r1) / tt) / 5
    else:
        gg = (sum(r1) / tt + b[3] + sum(r2) / kk + sum(r3) / j+sum(r4) / uu) / 5

    return gg
def sampleH(r11, mm1, mm2, mm3, mm4, sum1M, RM, CM,FM):
    r0=[]
    r1=[]
    r2 = []
    r3 = []
    r4=[]
    tt=0
    kk = 0
    j = 0
    uu = 0
    for i in range(0,len(r11)):
        if len(r11[i])==0:
            aa = b[4]
            r0.append(aa)

        if len(r11[i])==1:
            bb = find(r11[i], mm1, sum1M)
            r1.append(bb)
            tt=tt+1
        if len(r11[i])==2:
            cc = find(r11[i], mm2, RM)
            r2.append(cc)
            kk=kk+1
        if len(r11[i]) == 3:
            dd = find(r11[i], mm3, CM)
            r3.append(dd)
            j = j + 1
        if len(r11[i]) == 4:
            ee = find(r11[i], mm4, FM)
            r4.append(ee)
            uu = uu + 1
    if len(r1) == 0 and len(r2) == 0 and len(r3) == 0 and len(r4) == 0 and len(r0) == 0:
        gg = 0
    elif len(r1)==0 and len(r2) == 0 and len(r3) == 0 and len(r4) == 0:
        gg = (b[4]) / 5
    elif len(r0) == 0 and len(r2) == 0 and len(r3) == 0 and len(r4) == 0:
        gg = (sum(r1) / tt) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r3) == 0 and len(r4) == 0:
        gg = (sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) == 0 and len(r4) == 0:
        gg = (sum(r3) / j) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu) / 5

    elif len(r2) == 0 and len(r3) == 0 and len(r0) != 0 and len(r1) != 0 and len(r4) == 0:
        gg = (sum(r1) / tt + b[4]) / 5
    elif len(r0) == 0 and len(r1) != 0 and len(r2) != 0 and len(r3) == 0 and len(r4) == 0:
        gg = (sum(r1) / tt + sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r1) != 0 and len(r2) == 0 and len(r3) != 0 and len(r4) == 0:
        gg = (sum(r1) / tt + sum(r3) / j) / 5
    elif len(r0) != 0 and len(r1) == 0 and len(r2) != 0 and len(r3) == 0 and len(r4) == 0:
        gg = (b[4] + sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) != 0 and len(r3) != 0 and len(r4) == 0:
        gg = (sum(r3) / j + sum(r2) / kk) / 5
    elif len(r0) != 0 and len(r1) == 0 and len(r2) == 0 and len(r3) != 0 and len(r4) == 0:
        gg = (sum(r3) / j + b[4]) / 5

    elif len(r0) != 0 and len(r1) == 0 and len(r2) ==  0 and len(r3) ==  0 and len(r4) != 0:
        gg = (sum(r4) / uu + b[4]) / 5
    elif len(r0) == 0 and len(r1) != 0 and len(r2) ==  0 and len(r3) ==  0 and len(r4) != 0:
        gg = (sum(r4) / uu + sum(r1) / tt) / 5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) !=  0 and len(r3) ==  0 and len(r4) != 0:
        gg = (sum(r2) / kk + sum(r4) / uu) /5
    elif len(r0) == 0 and len(r1) == 0 and len(r2) ==  0 and len(r3) !=  0 and len(r4) != 0:
        gg = (sum(r3) / j + sum(r4) / uu) / 5

    elif len(r2) == 0 and len(r4) == 0 :
        gg = (sum(r1) / tt + b[4] + sum(r3) / j) / 5

    elif len(r3) == 0 and len(r4) == 0:
        gg = (sum(r1) / tt + b[4] + sum(r2) / kk) / 5

    elif len(r1) == 0 and len(r4) == 0:
        gg = (sum(r3) / j + b[4] + sum(r2) / kk) / 5

    elif len(r0) == 0 and len(r4) == 0:
        gg = (sum(r3) / j + sum(r1) / tt  + sum(r2) / kk) / 5
    elif len(r2) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt  + b[4]) / 5
    elif len(r1) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk  + b[4]) / 5
    elif len(r1) == 0 and len(r2) == 0:
        gg = (sum(r4) / uu + sum(r3) / j  + b[4]) / 5
    elif len(r0) == 0 and len(r3) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt + sum(r2) / kk) / 5
    elif len(r0) == 0 and len(r2) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt + sum(r3) / j) / 5
    elif len(r0) == 0 and len(r1) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + sum(r3) / j) / 5
    elif len(r0) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + sum(r3) / j + sum(r1) / tt) / 5
    elif len(r1) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + sum(r3) / j + b[4]) / 5
    elif len(r2) == 0:
        gg = (sum(r4) / uu + sum(r1) / tt + sum(r3) / j + b[4]) / 5
    elif len(r3) == 0:
        gg = (sum(r4) / uu + sum(r2) / kk + b[4] + sum(r1) / tt) / 5
    elif len(r4) == 0:
        gg = (sum(r3) / j + sum(r2) / kk + b[4] + sum(r1) / tt) / 5
    else:
        gg = (sum(r1) / tt + b[4] + sum(r2) / kk + sum(r3) / j+sum(r4) / uu) /5

    return gg
def find(index,mm3,R):

    h = 0
    for tt in range(0, len(mm3)):
        h = h + 1
        index=''.join(re.findall(r'[A-Za-z]', str(index)))

        if index == mm3[tt]:
            break
    r1 = R[h-1]
    return r1
def findee(r11):
    VSM=[]
    VSD=[]
    VSR=[]
    VSC=[]
    VSH=[]
    for tt in range(0, len(r11)):
        aa = r11[tt]
        str_listM = list(aa)
        str_listM.insert(0, 'M')
        wantM = ''.join(re.findall(r'[A-Za-z]', str(str_listM)))
        VSM.append(wantM)
        VSM.append(aa)
        str_listD= list(aa)
        str_listD.insert(0, 'D')
        wantD = ''.join(re.findall(r'[A-Za-z]', str(str_listD)))
        VSD.append(wantD)
        VSD.append(aa)
        str_listR = list(aa)
        str_listR.insert(0, 'R')
        wantR = ''.join(re.findall(r'[A-Za-z]', str(str_listR)))
        VSR.append(wantR)
        VSR.append(aa)
        str_listC = list(aa)
        str_listC.insert(0, 'C')
        wantC = ''.join(re.findall(r'[A-Za-z]', str(str_listC)))
        VSC.append(wantC)
        VSC.append(aa)
        str_listH= list(aa)
        str_listH.insert(0, 'H')
        wantH= ''.join(re.findall(r'[A-Za-z]', str(str_listH)))
        VSH.append(wantH)
        VSH.append(aa)
    eeM = list(set(VSM))
    eeR = list(set(VSR))
    eeD = list(set(VSD))
    eeC = list(set(VSC))
    eeH = list(set(VSH))
    tt = eeM + eeR + eeD + eeC+eeH
    eet = list(set(tt))
    s = []

    for i in range(len(eet)):
        for j in range(i + 1, len(eet)):
            if (sorted(eet[i]) == sorted(eet[j])):
                s.append(eet[j])
    s = list(set(s))
    if s != []:
        for jj in range(len(s)):
            if s[jj] not in eet:
                break
            eet.remove(s[jj])

    return len(eet)


def main():
    VS=["","M","R","D", "C", "H", "MM", "MD", "MC","MH", "DD", "DC","DH", "CC",
          "CH","HH","MR", "DR", "RR", "RC","RH",
     "MMM", "MMD", "MMR", "MMC", "MMH","MDD", "MDR", "MDC","MDH", "MRR",
    "MRC", "MRH","MCC","MCH","MHH","DDD", "DDR", "DDC", "DDH","DRR",
    "DRC","DRH", "DCC","DCH","DHH",
          "RRC", "RRH","RRR", "RCC","RCH","RHH", "CCC","CCH","CHH","HHH",
          "MMMM", "MMMD", "MMMR", "MMMC", "MMMH", "MMDD", "MMDR", "MMDC", "MMDH", "MMRR", "MMRC", "MMRH", "MMCC",
          "MMCH", "MMHH", "MDDD", "MDDR", "MDDC", "MDDH", "MDRR", "MDRC", "MDRH", "MDCC", "MDCH", "MDHH",
          "MRRC", "MRRH", "MRRR", "MRCC", "MRCH", "MRHH", "MCCC", "MCCH", "MCHH", "MHHH",
          "DDDD","DDDR","DDDC","DDDH","DDRR","DDRC","DDRH","DDCC","DDCH","DDHH","DRRR","DRRC","DRRH",
          "DRCC","DRCH","DRHH","DCCC","DCCH","DCHH","DHHH","RRRR","RRRC","RRRH","RRCC","RRCH","RRHH",
          "RCCC","RCCH","RCHH","RHHH","CCCC","CCCH","CCHH","CHHH","HHHH"]
    test = 2
    # 1，10,30

    x = 30
    workbook = xlwt.Workbook(encoding='utf-8')

    wbsheet1 = openexcel(workbook, "MSE")
    wbsheet2 = openexcel(workbook, "Spearman")
    wbsheet3 = openexcel(workbook, "ExpermentN")
    for ggg in range(0, 50):

        sum1M = [0] * 5
        sum1D = [0] * 5
        sum1R = [0] * 5
        sum1C = [0] * 5
        sum1H = [0] * 5
        RM = [0] * 15
        CM = [0] * 35
        FM = [0] * 70
        RD = [0] * 15
        CD = [0] * 35
        FD = [0] * 70
        RR = [0] * 15
        CR = [0] * 35
        FR = [0] * 70
        RC = [0] * 15
        CC = [0] * 35
        FC = [0] * 70
        RH= [0] * 15
        CH = [0] * 35
        FH = [0] * 70
        mm4 = []
        mm3=[]
        mm2=[]
        mm1 = []
        VSM=[]
        VSD = []
        VSR = []
        VSC = []
        VSH=[]
        f = 0
        QQ, WW = randoms(x)
        # QQ=["ew"]
        # WW=[12]

        rrM = 0
        ccM = 0
        ddM = 0
        ffM=0

        for i in range(0,len(VS)):
            want = ''.join(re.findall(r'[A-Za-z]', str(VS[i])))
            str_listM=list(want)
            str_listM.insert(0,'M')
            wantM = ''.join(re.findall(r'[A-Za-z]', str(str_listM)))
            VSM.append(wantM)
            str_listD = list(want)
            str_listD.insert(0, 'D')
            wantD = ''.join(re.findall(r'[A-Za-z]', str(str_listD)))
            VSD.append(wantD)
            str_listR = list(want)
            str_listR.insert(0, 'R')
            wantR = ''.join(re.findall(r'[A-Za-z]', str(str_listR)))
            VSR.append(wantR)
            str_listC = list(want)
            str_listC.insert(0, 'C')
            wantC = ''.join(re.findall(r'[A-Za-z]', str(str_listC)))
            VSC.append(wantC)
            str_listH = list(want)
            str_listH.insert(0, 'H')
            wantH = ''.join(re.findall(r'[A-Za-z]', str(str_listH)))
            VSH.append(wantH)
            sumM = findvalue(VSM[i],test) - findvalue(VS[i],test)
            sumD = findvalueD(VSD[i],test) - findvalue(VS[i],test)
            sumR = findvalueR(VSR[i],test) - findvalue(VS[i],test)
            sumC = findvalueC(VSC[i],test) - findvalue(VS[i],test)
            sumH = findvalueH(VSH[i], test) - findvalue(VS[i], test)
            for k in range(len(QQ)):

                if (VSM[i] == QQ[k]) and (VS[i] != QQ[k]):
                    sumM = WW[k] - findvalue(VS[i],test)
                if (VSM[i] == QQ[k]) and (VS[i] != QQ[k]):
                    sumM = findvalue(VSM[i],test) - WW[k]
                if (VSM[i] == QQ[k]) and (VS[i] == QQ[k]):
                    sumM = WW[k] - findvalue(VS[i],test)
                if (VSD[i] == QQ[k]) and (VS[i] != QQ[k]):
                    sumD = WW[k] - findvalue(VS[i],test)
                if (VSD[i] == QQ[k]) and (VS[i] != QQ[k]):
                    sumD = findvalue(VSD[i],test) - WW[k]
                if (VSD[i] == QQ[k]) and (VS[i] == QQ[k]):
                    sumD = WW[k] - findvalue(VS[i],test)
                if (VSR[i] == QQ[k]) and (VS[i] != QQ[k]):
                    sumR = WW[k] - findvalue(VS[i],test)
                if (VSR[i] == QQ[k]) and (VS[i] != QQ[k]):
                    sumR = findvalue(VSR[i],test) - WW[k]
                if (VSR[i] == QQ[k]) and (VS[i] == QQ[k]):
                    sumR = WW[k] - findvalue(VS[i],test)
                if (VSC[i] == QQ[k]) and (VS[i] != QQ[k]):
                    sumC = WW[k] - findvalue(VS[i],test)
                if (VSC[i] == QQ[k]) and (VS[i] != QQ[k]):
                    sumC = findvalue(VSC[i],test) - WW[k]
                if (VSC[i] == QQ[k]) and (VS[i] == QQ[k]):
                    sumC = WW[k] - findvalue(VS[i],test)
                if (VSH[i] == QQ[k]) and (VS[i] != QQ[k]):
                    sumH = WW[k] - findvalue(VS[i], test)
                if (VSH[i] == QQ[k]) and (VS[i] != QQ[k]):
                    sumH = findvalue(VSH[i], test) - WW[k]
                if (VSH[i] == QQ[k]) and (VS[i] == QQ[k]):
                    sumH = WW[k] - findvalue(VS[i], test)
            if len(VS[i]) == 0:
                SIM = sumM
                SID = sumD
                SIR = sumR
                SIC = sumC
                SIH = sumH
            if len(VS[i]) == 1:
                sum1M[ddM] = sumM
                sum1D[ddM] = sumD
                sum1R[ddM] = sumR
                sum1C[ddM] = sumC
                sum1H[ddM] = sumH
                ddM = ddM + 1
                mm1.append(VS[i])
            if len(VS[i]) == 2:
                RM[rrM] = sumM
                RD[rrM] = sumD
                RR[rrM] = sumR
                RC[rrM] = sumC
                RH[rrM] = sumH
                rrM = rrM + 1
                mm2.append(VS[i])

            if len(VS[i]) == 3:
                CM[ccM] = sumM
                CD[ccM] = sumD
                CR[ccM] = sumR
                CC[ccM] = sumC
                CH[ccM] = sumH
                ccM = ccM + 1
                mm3.append(VS[i])
            if len(VS[i]) == 4:
                FM[ffM] = sumM
                FD[ffM] = sumD
                FR[ffM] = sumR
                FC[ffM] = sumC
                FH[ffM] = sumH
                ffM = ffM + 1
                mm4.append(VS[i])
        sum1M1 = sum(sum1M) / 4
        rm = sum(RM) / 15
        cm = sum(CM) / 35
        fm = sum(FM) / 70
        sum1D1 = sum(sum1D) / 4
        rd = sum(RD) / 15
        cd = sum(CD) / 35
        fd = sum(FD)/70
        sum1R1 = sum(sum1R) / 4
        rr = sum(RR) / 15
        cr = sum(CR) / 35
        fr = sum(FR) / 70
        sum1C1 = sum(sum1C) / 4
        rc = sum(RC) / 15
        cc = sum(CC) / 35
        fc = sum(FC) / 70
        sum1H1 = sum(sum1H) / 4
        rh = sum(RH) / 15
        ch = sum(CH) / 35
        fh = sum(FH) / 70
        for k in range(len(QQ)):

            if ("M" == QQ[k]) :
                b[0]   = WW[k]
            elif ("D" == QQ[k]):
                b[1]=WW[k]
            elif ("R" == QQ[k]):
                b[2] = WW[k]
            elif ("C" == QQ[k]):
                b[3]=WW[k]
            elif ("H" == QQ[k]):
                b[4] = WW[k]
        print('M orginal shapley:')
        MM1=(sum1M1 + b[0] + rm + cm+fm) / 5
        print(MM1)
        print('D orginal shapley:')
        DD1=(sum1D1 + b[1] + rd + cd+fd) / 5
        print(DD1)
        print('R orginal shapley:')
        RR1=(sum1R1 + b[2] + rr + cr+fr) / 5
        print(RR1)
        print('C orginal shapley:')
        CC1=(sum1C1 + b[3] + rc + cc+fc) / 5
        print(CC1)
        print('H orginal shapley:')
        HH1 = (sum1H1 + b[4] + rh + ch + fh) / 5
        print(HH1)
        RS=[MM1,DD1,RR1,CC1,HH1]
        MSE11 = [0] * 50
        MSE22 = [0] * 50
        MSE33 = [0] * 50
        MSE44 = [0] * 50
        MSE55 = [0] * 50
        MSE66 = [0] * 50
        MSE77 = [0] * 50
        MSE88 = [0] * 50
        MSE99 = [0] * 50
        MSE10 = [0] * 50
        E1 = [0] * 50
        E2 = [0] * 50
        E3 = [0] * 50
        E4 = [0] * 50
        E5 = [0] * 50
        E6 = [0] * 50
        E7 = [0] * 50
        E8 = [0] * 50
        E9 = [0] * 50
        E10 = [0] * 50
        Sr11 = [0] * 50
        Sr22 = [0] * 50
        Sr33 = [0] * 50
        Sr44 = [0] * 50
        Sr55 = [0] * 50
        Sr66 = [0] * 50
        Sr77 = [0] * 50
        Sr88 = [0] * 50
        Sr99 = [0] * 50
        Sr10 = [0] * 50

        am = [SIM]
        am.extend(sum1M)
        am.extend(CM)
        am.extend(RM)
        ad = [SID]
        ad.extend(sum1D)
        ad.extend(CD)
        ad.extend(RD)
        ar = [SIR]
        ar.extend(sum1R)
        ar.extend(CR)
        ar.extend(RR)

        ac = [SIC]
        ac.extend(sum1C)
        ac.extend(CC)
        ac.extend(RC)

        ah = [SIH]
        ah.extend(sum1H)
        ah.extend(CH)
        ah.extend(RH)
        for ii in range(1, 51):
            # # select 1
            ii = ii - 1
            r11=random.sample(VS, 1)
            S1M = sampleM(r11, mm1, mm2, mm3, mm4, sum1M, RM, CM,FM)
            S1D = sampleD(r11, mm1, mm2, mm3, mm4, sum1D, RD, CD, FD)
            S1R = sampleR(r11, mm1, mm2, mm3, mm4, sum1R, RR, CR, FR)
            S1C = sampleC(r11, mm1, mm2, mm3, mm4, sum1C, RC, CC, FC)
            S1H = sampleH(r11, mm1, mm2, mm3, mm4, sum1H, RH, CH, FH)
            #MSE
            MSE11[ii] = mean_squared_error(RS, [S1M, S1D, S1R, S1C,S1H])
            #Spearman
            X1 = pd.Series(RS)
            Y1=pd.Series([S1M, S1D, S1R, S1C,S1H])
            Sr11[ii]= X1.corr(Y1, method="spearman")
            #实验数
            E1[ii]=findee(r11)


            # # select 2
            r11 = random.sample(VS, 3)
            S2M = sampleM(r11, mm1, mm2, mm3, mm4, sum1M, RM, CM,FM)
            S2D = sampleD(r11, mm1, mm2, mm3, mm4, sum1D, RD, CD, FD)
            S2R = sampleR(r11, mm1, mm2, mm3, mm4, sum1R, RR, CR, FR)
            S2C = sampleC(r11, mm1, mm2, mm3, mm4, sum1C, RC, CC, FC)
            S2H = sampleH(r11, mm1, mm2, mm3, mm4, sum1H, RH, CH, FH)
            # MSE
            MSE22[ii] = mean_squared_error(RS, [S2M, S2D, S2R, S2C,S2H])
            # Spearman
            X1 = pd.Series(RS)
            Y1 = pd.Series([S2M, S2D, S2R, S2C,S2H])
            Sr22[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            E2[ii] = findee(r11)




            # # select 3
            r11 = random.sample(VS, 5)
            S3M = sampleM(r11, mm1, mm2, mm3, mm4, sum1M, RM, CM, FM)
            S3D = sampleD(r11, mm1, mm2, mm3, mm4, sum1D, RD, CD, FD)
            S3R = sampleR(r11, mm1, mm2, mm3, mm4, sum1R, RR, CR, FR)
            S3C = sampleC(r11, mm1, mm2, mm3, mm4, sum1C, RC, CC, FC)
            S3H = sampleH(r11, mm1, mm2, mm3, mm4, sum1H, RH, CH, FH)
            # MSE
            MSE33[ii] = mean_squared_error(RS, [S3M, S3D, S3R, S3C,S3H])
            # Spearman
            Y1 = pd.Series([S3M, S3D, S3R, S3C,S3H])
            Sr33[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            E3[ii]= findee(r11)

            # select 4
            r11 = random.sample(VS, 7)
            S4M = sampleM(r11, mm1, mm2, mm3, mm4, sum1M, RM, CM, FM)
            S4D = sampleD(r11, mm1, mm2, mm3, mm4, sum1D, RD, CD, FD)
            S4R = sampleR(r11, mm1, mm2, mm3, mm4, sum1R, RR, CR, FR)
            S4C = sampleC(r11, mm1, mm2, mm3, mm4, sum1C, RC, CC, FC)
            S4H = sampleH(r11, mm1, mm2, mm3, mm4, sum1H, RH, CH, FH)
            MSE44[ii] = mean_squared_error(RS, [S4M, S4D, S4R, S4C,S4H])
            # Spearman
            Y1 = pd.Series([S4M, S4D, S4R, S4C,S4H])
            Sr44[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            E4[ii] = findee(r11)

            # # select 5
            r11 = random.sample(VS, 9)
            S5M = sampleM(r11, mm1, mm2, mm3, mm4, sum1M, RM, CM, FM)
            S5D = sampleD(r11, mm1, mm2, mm3, mm4, sum1D, RD, CD, FD)
            S5R = sampleR(r11, mm1, mm2, mm3, mm4, sum1R, RR, CR, FR)
            S5C = sampleC(r11, mm1, mm2, mm3, mm4, sum1C, RC, CC, FC)
            S5H = sampleH(r11, mm1, mm2, mm3, mm4, sum1H, RH, CH, FH)
            MSE55[ii] = mean_squared_error(RS, [S5M, S5D, S5R, S5C,S5H])
            # Spearman
            Y1 = pd.Series([S5M, S5D, S5R, S5C,S5H])
            Sr55[ii] = X1.corr(Y1, method="spearman")
            E5[ii] = findee(r11)

            # # select 6
            r11 = random.sample(VS, 17)
            S6M = sampleM(r11, mm1, mm2, mm3, mm4, sum1M, RM, CM, FM)
            S6D = sampleD(r11, mm1, mm2, mm3, mm4, sum1D, RD, CD, FD)
            S6R = sampleR(r11, mm1, mm2, mm3, mm4, sum1R, RR, CR, FR)
            S6C = sampleC(r11, mm1, mm2, mm3, mm4, sum1C, RC, CC, FC)
            S6H = sampleH(r11, mm1, mm2, mm3, mm4, sum1H, RH, CH, FH)
            MSE66[ii] = mean_squared_error(RS, [S6M, S6D, S6R, S6C,S6H])
            # Spearman
            Y1 = pd.Series([S6M, S6D, S6R, S6C,S6H])
            Sr66[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            E6[ii] = findee(r11)

            # # select 7
            r11 = random.sample(VS, 28)
            # r11 = random.sample(VS, 60)
            S7M = sampleM(r11, mm1, mm2, mm3, mm4, sum1M, RM, CM, FM)
            S7D = sampleD(r11, mm1, mm2, mm3, mm4, sum1D, RD, CD, FD)
            S7R = sampleR(r11, mm1, mm2, mm3, mm4, sum1R, RR, CR, FR)
            S7C = sampleC(r11, mm1, mm2, mm3, mm4, sum1C, RC, CC, FC)
            S7H = sampleH(r11, mm1, mm2, mm3, mm4, sum1H, RH, CH, FH)
            MSE77[ii] = mean_squared_error(RS, [S7M, S7D, S7R, S7C,S7H])
            # Spearman
            Y1 = pd.Series([S7M, S7D, S7R, S7C,S7H])
            Sr77[ii]= X1.corr(Y1, method="spearman")
            # 实验数
            E7[ii] = findee(r11)

            # # select 8
            r11 = random.sample(VS, 41)
            S8M = sampleM(r11, mm1, mm2, mm3, mm4, sum1M, RM, CM, FM)
            S8D = sampleD(r11, mm1, mm2, mm3, mm4, sum1D, RD, CD, FD)
            S8R = sampleR(r11, mm1, mm2, mm3, mm4, sum1R, RR, CR, FR)
            S8C = sampleC(r11, mm1, mm2, mm3, mm4, sum1C, RC, CC, FC)
            S8H = sampleH(r11, mm1, mm2, mm3, mm4, sum1H, RH, CH, FH)
            MSE88[ii] = mean_squared_error(RS, [S8M, S8D, S8R, S8C,S8H])
            # Spearman
            Y1 = pd.Series([S8M, S8D, S8R, S8C,S8H])
            Sr88[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            E8[ii] = findee(r11)

            # # select 9
            r11 = random.sample(VS, 60)
            S9M = sampleM(r11, mm1, mm2, mm3, mm4, sum1M, RM, CM, FM)
            S9D = sampleD(r11, mm1, mm2, mm3, mm4, sum1D, RD, CD, FD)
            S9R = sampleR(r11, mm1, mm2, mm3, mm4, sum1R, RR, CR, FR)
            S9C = sampleC(r11, mm1, mm2, mm3, mm4, sum1C, RC, CC, FC)
            S9H = sampleH(r11, mm1, mm2, mm3, mm4, sum1H, RH, CH, FH)
            MSE99[ii] = mean_squared_error(RS, [S9M, S9D, S9R, S9C,S9H])
            # Spearman
            Y1 = pd.Series([S9M, S9D, S9R, S9C,S9H])
            Sr99[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            E9[ii] = findee(r11)

            # # select 10

            S10M = sampleM(VS, mm1, mm2, mm3, mm4, sum1M, RM, CM, FM)
            S10D = sampleD(VS, mm1, mm2, mm3, mm4, sum1D, RD, CD, FD)
            S10R = sampleR(VS, mm1, mm2, mm3, mm4, sum1R, RR, CR, FR)
            S10C = sampleC(VS, mm1, mm2, mm3, mm4, sum1C, RC, CC, FC)
            S10H = sampleH(VS, mm1, mm2, mm3, mm4, sum1H, RH, CH, FH)
            MSE10[ii] = mean_squared_error(RS, [S10M, S10D, S10R, S10C,S10H])
            # Spearman
            Y1 = pd.Series([S10M, S10D, S10R, S10C,S10H])
            Sr10[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            E10[ii] = findee(VS)
            MSEG = [MSE11[ii], MSE22[ii], MSE33[ii], MSE44[ii], MSE55[ii], MSE66[ii], MSE77[ii], MSE88[ii], MSE99[ii],
                    MSE10[ii]]
            SMG = [Sr11[ii], Sr22[ii], Sr33[ii], Sr44[ii], Sr55[ii], Sr66[ii], Sr77[ii], Sr88[ii], Sr99[ii], Sr10[ii]]
            ENN = [E1[ii], E2[ii], E3[ii], E4[ii], E5[ii], E6[ii], E7[ii], E8[ii], E9[ii], E10[ii]]

            kk = 0

            for u in range(0, 10):
                wbsheet1.write(ii + (ggg * 50), u, MSEG[kk])
                wbsheet2.write(ii + (ggg * 50), u, SMG[kk])
                wbsheet3.write(ii + (ggg * 50), u, ENN[kk])
                kk = kk + 1

        print("it the " + ggg.__str__() + " run")
        workbook.save('Excel_testR53.xlsx')


























if __name__ == '__main__':
    main()