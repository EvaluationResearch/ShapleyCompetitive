#!/usr/bin/env python
from itertools import combinations
import xlrd
import xlwt
import re
import random
import numpy as np
import pandas as pd
from sklearn.metrics import mean_squared_error
a=["M","D","R","C"];
b=[0.82487932,0.11479157,0.60115175,0.39437373]#1
# b=[0.061912324,0.41768332,0.89046939,0.705243437]#2
# b=[0.111161148,0.387363514,0.945982619,0.536848777]#3
# b=[0.72883506, 0.14086356, 0.9985798,  0.43531148] #4
# b=[0.46712647,0.67700473,0.91128749,0.20564252]#3

def randoms(x):
    test=["M", "R", "D", "C", "MM", "MD", "MC", "DD", "DC", "CC", "MR", "DR", "RR", "RC",
    "MMM", "MMD", "MMR", "MMC", "MDD", "MDR", "MDC", "MRR", "MRC", "MCC",
    "DDD", "DDR", "DDC", "DRR", "DRC", "DCC", "RRC", "RRR", "RCC", "CCC"]

    F=random.sample(test, x)
    P = np.random.uniform(low=0, high=1, size=x)
    return F,P
def openexcel(workbook,lab):
    wb1 = workbook.add_sheet(lab, cell_overwrite_ok=True)

    return wb1

def findvalue(sub,test):
    if sub=='':
        return 0
    table = xlrd.open_workbook("16or1.xlsx")
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

    table = xlrd.open_workbook("test1.xlsx")
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

    table = xlrd.open_workbook("test1.xlsx")
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
    table = xlrd.open_workbook("test1.xlsx")
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


def sampleM(r11,mm1,mm2,mm3,sum1M,RM,CM):
    r0=[]
    r1=[]
    r2 = []
    r3 = []
    tt=0
    kk = 0
    j = 0
    for i in range(0,len(r11)):
        if len(r11[i])==0:
            cc = b[0]
            r0.append(cc)

        if len(r11[i])==1:
            cc = find(r11[i], mm1, sum1M)
            r1.append(cc)
            tt=tt+1
        if len(r11[i])==2:
            cc = find(r11[i], mm2, RM)
            r2.append(cc)
            kk=kk+1
        if len(r11[i]) == 3:
            dd = find(r11[i], mm3, CM)
            r3.append(dd)
            j = j + 1
    if len(r1)==0 and len(r2) == 0 and len(r3) == 0:
        gg = (b[0]) / 4
    elif len(r0) == 0 and len(r2) == 0 and len(r3) == 0:
        gg = (sum(r1) / tt) / 4
    elif len(r0) == 0 and len(r1) == 0 and len(r3) == 0:
        gg = (sum(r2) / kk) / 4
    elif len(r0) == 0 and len(r1) == 0 and len(r2) == 0:
        gg = (sum(r3) / j) / 4

    elif len(r2) == 0 and len(r3) == 0 and len(r0) != 0 and len(r1) != 0 :
        gg = (sum(r1) / tt + b[0]) / 4
    elif len(r0) == 0 and len(r1) != 0 and len(r2) != 0 and len(r3) == 0:
        gg = (sum(r1) / tt + sum(r2) / kk) / 4
    elif len(r0) == 0 and len(r1) != 0 and len(r2) == 0 and len(r3) != 0:
        gg = (sum(r1) / tt + sum(r3) / j) / 4
    elif len(r0) != 0 and len(r1) == 0 and len(r2) != 0 and len(r3) == 0:
        gg = (b[0] + sum(r2) / kk) / 4
    elif len(r0) == 0 and len(r1) == 0 and len(r2) != 0 and len(r3) != 0:
        gg = (sum(r3) / j + sum(r2) / kk) / 4
    elif len(r0) != 0 and len(r1) == 0 and len(r2) == 0 and len(r3) != 0:
        gg = (sum(r3) / j + b[0]) / 4

    elif len(r2) == 0:
        gg = (sum(r1) / tt + b[0] + sum(r3) / j) / 4

    elif len(r3) == 0:
        gg = (sum(r1) / tt + b[0] + sum(r2) / kk) / 4

    elif len(r1) == 0:
        gg = (sum(r3) / j + b[0] + sum(r2) / kk) / 4
    else:
        gg = (sum(r1) / tt + b[0] + sum(r2) / kk + sum(r3) / j) / 4

    return gg
def sampleD(r11,mm1,mm2,mm3,sum1M,RM,CM):
    r0 = []
    r1 = []
    r2 = []
    r3 = []
    tt = 0
    kk = 0
    j = 0
    for i in range(0, len(r11)):
        if len(r11[i]) == 0:
            cc = b[1]
            r0.append(cc)

        if len(r11[i]) == 1:
            cc = find(r11[i], mm1, sum1M)
            r1.append(cc)
            tt = tt + 1
        if len(r11[i]) == 2:
            cc = find(r11[i], mm2, RM)
            r2.append(cc)
            kk = kk + 1
        if len(r11[i]) == 3:
            dd = find(r11[i], mm3, CM)
            r3.append(dd)
            j = j + 1
    if len(r1) == 0 and len(r2) == 0 and len(r3) == 0:
        gg = (b[1]) / 4
    elif len(r0) == 0 and len(r2) == 0 and len(r3) == 0:
        gg = (sum(r1) / tt) / 4
    elif len(r0) == 0 and len(r1) == 0 and len(r3) == 0:
        gg = (sum(r2) / kk) / 4
    elif len(r0) == 0 and len(r1) == 0 and len(r2) == 0:
        gg = (sum(r3) / j) / 4

    elif len(r2) == 0 and len(r3) == 0 and len(r0) != 0 and len(r1) != 0:
        gg = (sum(r1) / tt + b[1]) / 4
    elif len(r0) == 0 and len(r1) != 0 and len(r2) != 0 and len(r3) == 0:
        gg = (sum(r1) / tt + sum(r2) / kk) / 4
    elif len(r0) == 0 and len(r1) != 0 and len(r2) == 0 and len(r3) != 0:
        gg = (sum(r1) / tt + sum(r3) / j) / 4

    elif len(r0) != 0 and len(r1) == 0 and len(r2) != 0 and len(r3) == 0:
        gg = (b[1] + sum(r2) / kk) / 4
    elif len(r0) == 0 and len(r1) == 0 and len(r2) != 0 and len(r3) != 0:
        gg = (sum(r3) / j + sum(r2) / kk) / 4
    elif len(r0) != 0 and len(r1) == 0 and len(r2) == 0 and len(r3) != 0:
        gg = (sum(r3) / j + b[1]) / 4

    elif len(r2) == 0:
        gg = (sum(r1) / tt + b[1] + sum(r3) / j) / 4

    elif len(r3) == 0:
        gg = (sum(r1) / tt + b[1] + sum(r2) / kk) / 4

    elif len(r1) == 0:
        gg = (sum(r3) / j + b[1] + sum(r2) / kk) / 4
    else:
        gg = (sum(r1) / tt + b[1] + sum(r2) / kk + sum(r3) / j) / 4

    return gg
def sampleR(r11,mm1,mm2,mm3,sum1M,RM,CM):
    r0 = []
    r1 = []
    r2 = []
    r3 = []
    tt = 0
    kk = 0
    j = 0
    for i in range(0, len(r11)):
        if len(r11[i]) == 0:
            cc = b[2]
            r0.append(cc)

        if len(r11[i]) == 1:
            cc = find(r11[i], mm1, sum1M)
            r1.append(cc)
            tt = tt + 1
        if len(r11[i]) == 2:
            cc = find(r11[i], mm2, RM)
            r2.append(cc)
            kk = kk + 1
        if len(r11[i]) == 3:
            dd = find(r11[i], mm3, CM)
            r3.append(dd)
            j = j + 1
    if len(r1) == 0 and len(r2) == 0 and len(r3) == 0:
        gg = (b[2]) / 4
    elif len(r0) == 0 and len(r2) == 0 and len(r3) == 0:
        gg = (sum(r1) / tt) / 4
    elif len(r0) == 0 and len(r1) == 0 and len(r3) == 0:
        gg = (sum(r2) / kk) / 4
    elif len(r0) == 0 and len(r1) == 0 and len(r2) == 0:
        gg = (sum(r3) / j) / 4

    elif len(r2) == 0 and len(r3) == 0 and len(r0) != 0 and len(r1) != 0:
        gg = (sum(r1) / tt + b[2]) / 4
    elif len(r0) == 0 and len(r1) != 0 and len(r2) != 0 and len(r3) == 0:
        gg = (sum(r1) / tt + sum(r2) / kk) / 4
    elif len(r0) == 0 and len(r1) != 0 and len(r2) == 0 and len(r3) != 0:
        gg = (sum(r1) / tt + sum(r3) / j) / 4
    elif len(r0) != 0 and len(r1) == 0 and len(r2) != 0 and len(r3) == 0:
        gg = (b[2] + sum(r2) / kk) / 4
    elif len(r0) == 0 and len(r1) == 0 and len(r2) != 0 and len(r3) != 0:
        gg = (sum(r3) / j + sum(r2) / kk) / 4
    elif len(r0) != 0 and len(r1) == 0 and len(r2) == 0 and len(r3) != 0:
        gg = (sum(r3) / j + b[2]) / 4

    elif len(r2) == 0:
        gg = (sum(r1) / tt + b[2] + sum(r3) / j) / 4

    elif len(r3) == 0:
        gg = (sum(r1) / tt + b[2] + sum(r2) / kk) / 4

    elif len(r1) == 0:
        gg = (sum(r3) / j + b[2] + sum(r2) / kk) / 4
    else:
        gg = (sum(r1) / tt + b[2] + sum(r2) / kk + sum(r3) / j) / 4
    return gg
def sampleC(r11,mm1,mm2,mm3,sum1M,RM,CM):
    r0 = []
    r1 = []
    r2 = []
    r3 = []
    tt = 0
    kk = 0
    j = 0
    for i in range(0, len(r11)):
        if len(r11[i]) == 0:
            cc = b[3]
            r0.append(cc)

        if len(r11[i]) == 1:
            cc = find(r11[i], mm1, sum1M)
            r1.append(cc)
            tt = tt + 1
        if len(r11[i]) == 2:
            cc = find(r11[i], mm2, RM)
            r2.append(cc)
            kk = kk + 1
        if len(r11[i]) == 3:
            dd = find(r11[i], mm3, CM)
            r3.append(dd)
            j = j + 1
    if len(r1) == 0 and len(r2) == 0 and len(r3) == 0:
        gg = (b[3]) / 4
    elif len(r0) == 0 and len(r2) == 0 and len(r3) == 0:
        gg = (sum(r1) / tt) / 4
    elif len(r0) == 0 and len(r1) == 0 and len(r3) == 0:
        gg = (sum(r2) / kk) / 4
    elif len(r0) == 0 and len(r1) == 0 and len(r2) == 0:
        gg = (sum(r3) / j) / 4

    elif len(r2) == 0 and len(r3) == 0 and len(r0) != 0 and len(r1) != 0:
        gg = (sum(r1) / tt + b[3]) / 4
    elif len(r0) == 0 and len(r1) != 0 and len(r2) != 0 and len(r3) == 0:
        gg = (sum(r1) / tt + sum(r2) / kk) / 4
    elif len(r0) == 0 and len(r1) != 0 and len(r2) == 0 and len(r3) != 0:
        gg = (sum(r1) / tt + sum(r3) / j) / 4
    elif len(r0) != 0 and len(r1) == 0 and len(r2) != 0 and len(r3) == 0:
        gg = (b[3] + sum(r2) / kk) / 4
    elif len(r0) == 0 and len(r1) == 0 and len(r2) != 0 and len(r3) != 0:
        gg = (sum(r3) / j + sum(r2) / kk) / 4
    elif len(r0) != 0 and len(r1) == 0 and len(r2) == 0 and len(r3) != 0:
        gg = (sum(r3) / j + b[3]) / 4

    elif len(r2) == 0:
        gg = (sum(r1) / tt + b[3] + sum(r3) / j) / 4

    elif len(r3) == 0:
        gg = (sum(r1) / tt + b[3] + sum(r2) / kk) / 4

    elif len(r1) == 0:
        gg = (sum(r3) / j + b[3] + sum(r2) / kk) / 4
    else:
        gg = (sum(r1) / tt + b[3] + sum(r2) / kk + sum(r3) / j) / 4
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
    VSM = []
    VSD = []
    VSR = []
    VSC = []

    for tt in range(0, len(r11)):
        aa = r11[tt]
        str_listM = list(aa)
        str_listM.insert(0, 'M')
        wantM = ''.join(re.findall(r'[A-Za-z]', str(str_listM)))
        VSM.append(wantM)
        VSM.append(aa)
        str_listD = list(aa)
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
    VSM.append('')
    eeM = list(set(VSM))
    eeR = list(set(VSR))
    eeD = list(set(VSD))
    eeC = list(set(VSC))
    tt = eeM + eeR + eeD + eeC
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
    VS=["","M","R","D", "C", "MM", "MD",  "MC", "DD",  "DC", "CC", "MR", "DR", "RR", "RC",
        "MMM", "MMD", "MMR", "MMC", "MDD", "MDR", "MDC", "MRR","MRC", "MCC",
        "DDD", "DDR", "DDC", "DRR", "DRC", "DCC", "RRC", "RRR", "RCC", "CCC"]
    test=1
    #1,5,15

    x=1
    workbook = xlwt.Workbook(encoding='utf-8')

    wbsheet1 = openexcel(workbook, "MSE")
    wbsheet2 = openexcel(workbook, "Spearman")
    wbsheet3 = openexcel(workbook, "ExpermentN")
    for ggg in range(0, 50):
        sum1M = [0] * 4
        sum1D = [0] * 4
        sum1R = [0] * 4
        sum1C =[0] * 4

        RM = [0] * 10
        CM = [0] * 20
        RD = [0] * 10
        CD = [0] * 20
        RR = [0] * 10
        CR = [0] * 20
        RC = [0] * 10
        CC = [0] * 20

        mm3 = []
        mm2 = []
        mm1=[]
        VSM = []
        VSD = []
        VSR = []
        VSC = []

        QQ, WW = randoms(x)
        # QQ = ["EW"]
        # WW = [4]

        rrM = 0
        ccM = 0
        ddM=0

        for i in range(0, len(VS)):
            want = ''.join(re.findall(r'[A-Za-z]', str(VS[i])))
            str_listM = list(want)
            str_listM.insert(0, 'M')
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
            # 0.111161148, 0.387363514, 0.945982619, 0.536848777
            sumM = findvalue(VSM[i], test) - findvalue(VS[i], test)
            sumD = findvalueD(VSD[i], test) - findvalue(VS[i], test)
            sumR = findvalueR(VSR[i], test) - findvalue(VS[i], test)
            sumC = findvalueC(VSC[i], test) - findvalue(VS[i], test)

            for k in range(len(QQ)):

                if (VSM[i] == QQ[k]) and (VS[i] != QQ[k]):
                    sumM = WW[k] - findvalue(VS[i], test)
                if (VSM[i] == QQ[k]) and (VS[i] != QQ[k]):
                    sumM = findvalue(VSM[i], test) - WW[k]
                if (VSM[i] == QQ[k]) and (VS[i] == QQ[k]):
                    sumM = WW[k] - findvalue(VS[i], test)
                if (VSD[i] == QQ[k]) and (VS[i] != QQ[k]):
                    sumD = WW[k] - findvalue(VS[i], test)
                if (VSD[i] == QQ[k]) and (VS[i] != QQ[k]):
                    sumD = findvalue(VSD[i], test) - WW[k]
                if (VSD[i] == QQ[k]) and (VS[i] == QQ[k]):
                    sumD = WW[k] - findvalue(VS[i], test)
                if (VSR[i] == QQ[k]) and (VS[i] != QQ[k]):
                    sumR = WW[k] - findvalue(VS[i], test)
                if (VSR[i] == QQ[k]) and (VS[i] != QQ[k]):
                    sumR = findvalue(VSR[i], test) - WW[k]
                if (VSR[i] == QQ[k]) and (VS[i] == QQ[k]):
                    sumR = WW[k] - findvalue(VS[i], test)
                if (VSC[i] == QQ[k]) and (VS[i] != QQ[k]):
                    sumC = WW[k] - findvalue(VS[i], test)
                if (VSC[i] == QQ[k]) and (VS[i] != QQ[k]):
                    sumC = findvalue(VSC[i], test) - WW[k]
                if (VSC[i] == QQ[k]) and (VS[i] == QQ[k]):
                    sumC = WW[k] - findvalue(VS[i], test)
            sumM = float(format(sumM, '.9f'))
            sumD = float(format(sumD, '.9f'))
            sumR = float(format(sumR, '.9f'))
            sumC = float(format(sumC, '.9f'))
            if len(VS[i]) == 0:
                SIM=sumM
                SID=sumD
                SIR=sumR
                SIC=sumC
            if len(VS[i]) == 1:
                sum1M[ddM] = sumM
                sum1D[ddM] =  sumD
                sum1R[ddM]= sumR
                sum1C[ddM]=  sumC
                ddM=ddM+1
                mm1.append(VS[i])
            if len(VS[i]) == 2:
                RM[rrM] = sumM
                RD[rrM] = sumD
                RR[rrM] = sumR
                RC[rrM] = sumC
                rrM = rrM + 1
                mm2.append(VS[i])

            if len(VS[i]) == 3:
                CM[ccM] = sumM
                CD[ccM] = sumD
                CR[ccM] = sumR
                CC[ccM] = sumC
                ccM = ccM + 1
                mm3.append(VS[i])
        sum1M1=  sum(sum1M) / 4
        rm = sum(RM) / 10
        rm = float(format(rm, '.9f'))
        cm = sum(CM) / 20
        cm = float(format(cm, '.9f'))
        sum1D1 = sum(sum1D) / 4
        rd = sum(RD) / 10
        rd = float(format(rd, '.9f'))
        cd = sum(CD) / 20
        cd = float(format(cd, '.9f'))
        sum1R1 = sum(sum1R) / 4
        rr = sum(RR) / 10
        rr = float(format(rr, '.9f'))
        cr = sum(CR) / 20
        cr = float(format(cr, '.9f'))
        sum1C1 = sum(sum1C) / 4
        rc = sum(RC) / 10
        rc = float(format(rc, '.9f'))
        cc = sum(CC) / 20
        cc = float(format(cc, '.9f'))
        for k in range(len(QQ)):

            if ("M" == QQ[k]):
                b[0] = WW[k]
            elif ("D" == QQ[k]):
                b[1] = WW[k]
            elif ("R" == QQ[k]):
                b[2] = WW[k]
            elif ("C" == QQ[k]):
                b[3] = WW[k]
        print('M orginal shapley:')
        MM1 = (sum1M1 + b[0] + rm + cm) / 4
        MM1 = float(format(MM1, '.9f'))
        print(MM1)
        print('D orginal shapley:')
        DD1 = (sum1D1 + b[1] + rd + cd) / 4
        DD1 = float(format(DD1, '.9f'))
        print(DD1)
        print('R orginal shapley:')
        RR1 = (sum1R1 + b[2] + rr + cr) / 4
        RR1 = float(format(RR1, '.9f'))
        print(RR1)
        print('C orginal shapley:')
        CC1 = (sum1C1 + b[3] + rc + cc) / 4
        CC1 = float(format(CC1, '.9f'))
        print(CC1)
        RS = [MM1, DD1, RR1, CC1]
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



        for ii in range(1, 51):
            ii = ii - 1
            # # select 1
            r11 = random.sample(VS, 1)
            S1M=  sampleM(r11, mm1, mm2, mm3, sum1M, RM, CM)
            S1D = sampleD(r11, mm1, mm2, mm3, sum1D, RD, CD)
            S1R = sampleR(r11, mm1, mm2, mm3, sum1R, RR, CR)
            S1C = sampleC(r11, mm1, mm2, mm3, sum1C, RC, CC)
            # MSE
            MSE11[ii] = mean_squared_error(RS, [S1M, S1D, S1R, S1C])

            # Spearman
            X1 = pd.Series(RS)
            Y1 = pd.Series([S1M, S1D, S1R, S1C])
            Sr11[ii] = X1.corr(Y1, method="spearman")

            # 实验数

            E1[ii] = findee(r11)+1

            # # select 2
            r11 = random.sample(VS, 2)
            S2M = sampleM(r11, mm1, mm2, mm3, sum1M, RM, CM)
            S2D = sampleD(r11, mm1, mm2, mm3, sum1D, RD, CD)
            S2R = sampleR(r11, mm1, mm2, mm3, sum1R, RR, CR)
            S2C = sampleC(r11, mm1, mm2, mm3, sum1C, RC, CC)
            # MSE
            MSE22[ii] = mean_squared_error(RS, [S2M, S2D, S2R, S2C])
            # Spearman
            X1 = pd.Series(RS)
            Y1 = pd.Series([S2M, S2D, S2R, S2C])
            Sr22[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            E2[ii] = findee(r11)

            # # select 3
            r11 = random.sample(VS, 3)
            S3M = sampleM(r11, mm1, mm2, mm3, sum1M, RM, CM)
            S3D = sampleD(r11, mm1, mm2, mm3, sum1D, RD, CD)
            S3R = sampleR(r11, mm1, mm2, mm3, sum1R, RR, CR)
            S3C = sampleC(r11, mm1, mm2, mm3, sum1C, RC, CC)
            # MSE
            MSE33[ii] = mean_squared_error(RS, [S3M, S3D, S3R, S3C])
            # Spearman
            X1 = pd.Series(RS)
            Y1 = pd.Series([S3M, S3D, S3R, S3C])
            Sr33[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            E3[ii] = findee(r11)

            # # select 4
            r11 = random.sample(VS, 4)
            S4M = sampleM(r11, mm1, mm2, mm3, sum1M, RM, CM)
            S4D = sampleD(r11, mm1, mm2, mm3, sum1D, RD, CD)
            S4R = sampleR(r11, mm1, mm2, mm3, sum1R, RR, CR)
            S4C = sampleC(r11, mm1, mm2, mm3, sum1C, RC, CC)
            # MSE
            MSE44[ii] = mean_squared_error(RS, [S4M, S4D, S4R, S4C])
            # Spearman
            X1 = pd.Series(RS)
            Y1 = pd.Series([S4M, S4D, S4R, S4C])
            Sr44[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            E4[ii] = findee(r11)

            # # select 5
            r11 = random.sample(VS, 5)
            S5M = sampleM(r11, mm1, mm2, mm3, sum1M, RM, CM)
            S5D = sampleD(r11, mm1, mm2, mm3, sum1D, RD, CD)
            S5R = sampleR(r11, mm1, mm2, mm3, sum1R, RR, CR)
            S5C = sampleC(r11, mm1, mm2, mm3, sum1C, RC, CC)
            # MSE
            MSE55[ii] = mean_squared_error(RS, [S5M, S5D, S5R, S5C])
            # Spearman
            Y1 = pd.Series([S5M, S5D, S5R, S5C])
            Sr55[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            E5[ii] = findee(r11)

            # select 6
            r11 = random.sample(VS, 7)
            S6M = sampleM(r11, mm1, mm2, mm3, sum1M, RM, CM)
            S6D = sampleD(r11, mm1, mm2, mm3, sum1D, RD, CD)
            S6R = sampleR(r11, mm1, mm2, mm3, sum1R, RR, CR)
            S6C = sampleC(r11, mm1, mm2, mm3, sum1C, RC, CC)
            MSE66[ii] = mean_squared_error(RS, [S6M, S6D, S6R, S6C])
            # Spearman
            Y1 = pd.Series([S6M, S6D, S6R, S6C])
            Sr66[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            E6[ii] = findee(r11)

            # # select 7
            r11 = random.sample(VS, 11)
            S7M = sampleM(r11, mm1, mm2, mm3, sum1M, RM, CM)
            S7D = sampleD(r11, mm1, mm2, mm3, sum1D, RD, CD)
            S7R = sampleR(r11, mm1, mm2, mm3, sum1R, RR, CR)
            S7C = sampleC(r11, mm1, mm2, mm3, sum1C, RC, CC)
            MSE77[ii] = mean_squared_error(RS, [S7M, S7D, S7R, S7C])
            # Spearman
            Y1 = pd.Series([S7M, S7D, S7R, S7C])
            Sr77[ii] = X1.corr(Y1, method="spearman")
            E7[ii] = findee(r11)

            # # select 8
            r11 = random.sample(VS, 16)
            S8M = sampleM(r11, mm1, mm2, mm3, sum1M, RM, CM)
            S8D = sampleD(r11, mm1, mm2, mm3, sum1D, RD, CD)
            S8R = sampleR(r11, mm1, mm2, mm3, sum1R, RR, CR)
            S8C = sampleC(r11, mm1, mm2, mm3, sum1C, RC, CC)
            MSE88[ii] = mean_squared_error(RS, [S8M, S8D, S8R, S8C])
            # Spearman
            Y1 = pd.Series([S8M, S8D, S8R, S8C])
            Sr88[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            E8[ii] = findee(r11)

            # # select 9
            r11 = random.sample(VS, 23)
            S9M = sampleM(r11, mm1, mm2, mm3, sum1M, RM, CM)
            S9D = sampleD(r11, mm1, mm2, mm3, sum1D, RD, CD)
            S9R = sampleR(r11, mm1, mm2, mm3, sum1R, RR, CR)
            S9C = sampleC(r11, mm1, mm2, mm3, sum1C, RC, CC)
            MSE99[ii] = mean_squared_error(RS, [S9M, S9D, S9R, S9C])
            # Spearman
            Y1 = pd.Series([S9M, S9D, S9R, S9C])
            Sr99[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            E9[ii] = findee(r11)

            # # select 10
            r11 = random.sample(VS, 35)
            S10M = sampleM(r11, mm1, mm2, mm3, sum1M, RM, CM)
            S10D = sampleD(r11, mm1, mm2, mm3, sum1D, RD, CD)
            S10R = sampleR(r11, mm1, mm2, mm3, sum1R, RR, CR)
            S10C = sampleC(r11, mm1, mm2, mm3, sum1C, RC, CC)
            MSE10[ii] = mean_squared_error(RS, [S10M, S10D, S10R, S10C])
            # Spearman
            Y1 = pd.Series([S10M, S10D, S10R, S10C])
            Sr10[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            E10[ii] = findee(r11)
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

        print("it the "+ggg.__str__()+" run")
    workbook.save('Excel_testR1.xlsx')

























if __name__ == '__main__':
    main()
