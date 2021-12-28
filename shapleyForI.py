#!/usr/bin/env python
from itertools import combinations
import xlrd
import xlwt
import re
import random
import numpy as np
import pandas as pd
from sklearn.metrics import mean_squared_error
import Levenshtein
import copy
a=["M","D","R","C"];
b=[0.82487932,0.11479157,0.60115175,0.39437373]#1
# b=[0.061912324,0.41768332,0.89046939,0.705243437]#2
# b=[0.111161148,0.387363514,0.945982619,0.536848777]#3
# b=[0.72883506, 0.14086356, 0.9985798,  0.43531148] #4
# b=[0.46712647,0.67700473,0.91128749,0.20564252]#5
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



def sampleI(mm,C,N):
    w = []
    w1=[]
    w2=[]
    w3=[]
    mnew=copy.deepcopy(mm)
    for kk in range(len(C)):
        if C[kk]==0:
            C.remove(0)
    for i in range(len(C)):
        if C[i] not in mm:
            break
        mnew.remove(C[i])

    L=len(C)
    for i in range(0, len(mnew)):
        if Levenshtein.distance(mnew[i], C[L-1]) == 3:
            w.append(mnew[i])
        if Levenshtein.distance(mnew[i], C[L-1]) == 2:
            w1.append(mnew[i])
        if Levenshtein.distance(mnew[i], C[L-1]) == 1:
            w2.append(mnew[i])
        if Levenshtein.distance(mnew[i], C[L-1]) == 0:
            w3.append(mnew[i])
    T1 = w + w1
    T2 = w + w1 + w2
    T3 = w + w1 + w2 + w3
    if len(w) > N or len(w) == N:
        b1 = random.sample(w, N)
        b1 = ''.join(re.findall(r'[A-Za-z]', str(b1)))
        return b1
    elif len(T1) > N or len(T1) == N:
        b2 = random.sample(T1, N)
        b2 = ''.join(re.findall(r'[A-Za-z]', str(b2)))
        return b2
    elif len(T2) > N or len(T2) == N:
        b3 = random.sample(T2, N)
        b3 = ''.join(re.findall(r'[A-Za-z]', str(b3)))
        return b3
    else:
        b4 = random.sample(T3, N)
        b4 = ''.join(re.findall(r'[A-Za-z]', str(b4)))
        return b4






def sampleIM(mm,C):
    if C=="DDD" or C=="MMM" or C=="RRR" or C=="CCC":
        C="MDR"
    w=[]
    for i in range(0, len(mm)):
        if Levenshtein.distance(mm[i], C) == 3:
            w.append(mm[i])
        if w!=[]:
            break
    if w==[]:
        for i in range(0, len(mm)):
            if Levenshtein.distance(mm[i], C) == 2:
                w.append(mm[i])
            if w != []:
                break


    return w[0]


def sampleII(mm, c11, N):
    for i in range(0, N):
        b = sampleI(mm, c11, 1)
        b = ''.join(re.findall(r'[A-Za-z]', str(b)))
        c11.append(b)
    return c11









def sampleM(r11,c11,mm2,mm3,RM,CM,sum1):
    r1=[]
    c1=[]
    for i in range(0,len(r11)):
        r1.append(find(r11[i], mm2, RM))
    for i in range(0,len(c11)):
        c1.append(find(c11[i], mm3, CM))
    if len(r1) == 0 and len(c1) == 0:
        gg = (sum1 + b[0]) / 4
    elif len(r1) == 0:
        gg = (sum1 + b[0] + sum(c1) / len(c1)) / 4

    elif len(c1) == 0:
        gg = (sum1 + b[0] + sum(r1) / len(r1)) / 4
    else:
        gg = (sum1 + b[0] + sum(r1) / len(r1) + sum(c1) / len(c1)) / 4

    return gg
def sampleD(r11,c11,mm2,mm3,RD,CD,sum1):
    r1 = []
    c1 = []
    for i in range(0, len(r11)):
        r1.append(find(r11[i], mm2, RD))
    for i in range(0, len(c11)):
        c1.append(find(c11[i], mm3, CD))
    if len(r1) == 0 and len(c1) == 0:
        gg = (sum1 + b[1]) / 4
    elif len(r1) == 0:
        gg = (sum1 + b[1] + sum(c1) / len(c1)) / 4

    elif len(c1) == 0:
        gg = (sum1 + b[1] + sum(r1) / len(r1)) / 4
    else:
        gg = (sum1 + b[1] + sum(r1) / len(r1) + sum(c1) / len(c1)) / 4
    return gg
def sampleR(r11,c11,mm2,mm3,RR,CR,sum1):
    r1 = []
    c1 = []
    for i in range(0, len(r11)):
        r1.append(find(r11[i], mm2, RR))
    for i in range(0, len(c11)):
        c1.append(find(c11[i], mm3, CR))
    if len(r1) == 0 and len(c1) == 0:
        gg = (sum1 + b[2]) / 4
    elif len(r1) == 0:
        gg = (sum1 + b[2] + sum(c1) / len(c1)) / 4

    elif len(c1) == 0:
        gg = (sum1 + b[2] + sum(r1) / len(r1)) / 4
    else:
        gg = (sum1 + b[2] + sum(r1) / len(r1) + sum(c1) / len(c1)) / 4
    return gg
def sampleC(r11,c11,mm2,mm3,RC,CC,sum1):
    r1 = []
    c1 = []
    for i in range(0, len(r11)):
        r1.append(find(r11[i], mm2, RC))
    for i in range(0, len(c11)):
        c1.append(find(c11[i], mm3, CC))
    if len(r1) == 0 and len(c1) == 0:
        gg = (sum1 + b[3]) / 4
    elif len(r1) == 0:
        gg = (sum1 + b[3] + sum(c1) / len(c1)) / 4

    elif len(c1) == 0:
        gg = (sum1 + b[3] + sum(r1) / len(r1)) / 4
    else:
        gg = (sum1 + b[3] + sum(r1) / len(r1) + sum(c1) / len(c1)) / 4
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

    test =1
    # 2,3,15,5,15

    x =1
    workbook = xlwt.Workbook(encoding='utf-8')

    wbsheet1 = openexcel(workbook, "MSE")
    wbsheet2 = openexcel(workbook, "Spearman")
    wbsheet3 = openexcel(workbook, "ExpermentN")
    for ggg in range(0, 50):
        sum1M = 0
        sum1D = 0
        sum1R = 0
        sum1C = 0

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
        VSM = []
        VSD = []
        VSR = []
        VSC = []

        QQ, WW = randoms(x)
        # QQ = ["EW"]
        # WW = [4]
        rrM = 0
        ccM = 0

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
            if len(VS[i]) == 1:
                sum1M = sum1M + 1 / 4 * sumM
                sum1D = sum1D + 1 / 4 * sumD
                sum1R = sum1R + 1 / 4 * sumR
                sum1C = sum1C + 1 / 4 * sumC

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

        rm = sum(RM) / 10
        cm = sum(CM) / 20
        rd = sum(RD) / 10
        cd = sum(CD) / 20
        rr = sum(RR) / 10
        cr = sum(CR) / 20
        rc = sum(RC) / 10
        cc = sum(CC) / 20
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
        MM1 = (sum1M + b[0] + rm + cm) / 4
        print(MM1)
        print('D orginal shapley:')
        DD1 = (sum1D + b[1] + rd + cd) / 4
        print(DD1)
        print('R orginal shapley:')
        RR1 = (sum1R + b[2] + rr + cr) / 4
        print(RR1)
        print('C orginal shapley:')
        CC1 = (sum1C + b[3] + rc + cc) / 4
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
            r11 =[] # R
            c11 = []  # C
            S1M = sampleM(r11, c11, mm2, mm3, RM, CM, sum1M)
            S1D = sampleD(r11, c11, mm2, mm3, RD, CD, sum1D)
            S1R = sampleR(r11, c11, mm2, mm3, RR, CR, sum1R)
            S1C = sampleC(r11, c11, mm2, mm3, RC, CC, sum1C)
            # MSE

            MSE11[ii] = mean_squared_error(RS, [S1M, S1D, S1R, S1C])

            # Spearman
            X1 = pd.Series(RS)
            Y1 = pd.Series([S1M, S1D, S1R, S1C])
            Sr11[ii] = X1.corr(Y1, method="spearman")

            # 实验数

            r11.extend(c11)
            E1[ii] = findee(r11) + 5

            # # select 2
            r11 = []
            a1 = random.sample(mm2, 1)
            a1 = ''.join(re.findall(r'[A-Za-z]', str(a1)))
            r11.append(a1)
            c11 = []
            S2M = sampleM(r11, c11, mm2, mm3, RM, CM, sum1M)
            S2D = sampleD(r11, c11, mm2, mm3, RD, CD, sum1D)
            S2R = sampleR(r11, c11, mm2, mm3, RR, CR, sum1R)
            S2C = sampleC(r11, c11, mm2, mm3, RC, CC, sum1C)
            # MSE
            MSE22[ii] = mean_squared_error(RS, [S2M, S2D, S2R, S2C])
            # Spearman
            X1 = pd.Series(RS)
            Y1 = pd.Series([S2M, S2D, S2R, S2C])
            Sr22[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            r11.extend(c11)
            E2[ii] = findee(r11) + 5

            # # select 3
            r11 = []  # 1
            c11 = []  # 1
            a1 = random.sample(mm2, 1)
            a1 = ''.join(re.findall(r'[A-Za-z]', str(a1)))
            r11.append(a1)
            b1 = random.sample(mm3, 1)
            b1 = ''.join(re.findall(r'[A-Za-z]', str(b1)))
            c11.append(b1)
            S3M = sampleM(r11, c11, mm2, mm3, RM, CM, sum1M)
            S3D = sampleD(r11, c11, mm2, mm3, RD, CD, sum1D)
            S3R = sampleR(r11, c11, mm2, mm3, RR, CR, sum1R)
            S3C = sampleC(r11, c11, mm2, mm3, RC, CC, sum1C)
            # MSE
            MSE33[ii] = mean_squared_error(RS, [S3M, S3D, S3R, S3C])
            # Spearman
            Y1 = pd.Series([S3M, S3D, S3R, S3C])
            Sr33[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            r11.extend(c11)
            E3[ii] = findee(r11) + 5

            # select 4
            r11 = []
            c11 = []
            a1 = random.sample(mm2, 1)
            a1 = ''.join(re.findall(r'[A-Za-z]', str(a1)))
            r11.append(a1)
            r11 = sampleII(mm2, r11,1)

            b1 = random.sample(mm3, 1)
            b1 = ''.join(re.findall(r'[A-Za-z]', str(b1)))
            c11.append(b1)
            S4M = sampleM(r11, c11, mm2, mm3, RM, CM, sum1M)
            S4D = sampleD(r11, c11, mm2, mm3, RD, CD, sum1D)
            S4R = sampleR(r11, c11, mm2, mm3, RR, CR, sum1R)
            S4C = sampleC(r11, c11, mm2, mm3, RC, CC, sum1C)
            MSE44[ii] = mean_squared_error(RS, [S4M, S4D, S4R, S4C])
            # Spearman
            Y1 = pd.Series([S4M, S4D, S4R, S4C])
            Sr44[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            r11.extend(c11)
            E4[ii] = findee(r11) + 5

            # # select 5
            r11 = []  # 2
            c11 = []  #2
            a1 = random.sample(mm2, 1)
            a1 = ''.join(re.findall(r'[A-Za-z]', str(a1)))
            r11.append(a1)
            r11 = sampleII(mm2, r11, 1)

            b1 = random.sample(mm3, 1)
            b1 = ''.join(re.findall(r'[A-Za-z]', str(b1)))
            c11.append(b1)
            c11 = sampleII(mm3, c11, 1)
            S5M = sampleM(r11, c11, mm2, mm3, RM, CM, sum1M)
            S5D = sampleD(r11, c11, mm2, mm3, RD, CD, sum1D)
            S5R = sampleR(r11, c11, mm2, mm3, RR, CR, sum1R)
            S5C = sampleC(r11, c11, mm2, mm3, RC, CC, sum1C)
            MSE55[ii] = mean_squared_error(RS, [S5M, S5D, S5R, S5C])
            # Spearman
            Y1 = pd.Series([S5M, S5D, S5R, S5C])
            Sr55[ii] = X1.corr(Y1, method="spearman")
            r11.extend(c11)
            E5[ii] = findee(r11) + 5

            # # select 6
            r11 = []  # 3
            c11 = []  # 2
            a1 = random.sample(mm2, 1)
            a1 = ''.join(re.findall(r'[A-Za-z]', str(a1)))
            r11.append(a1)
            r11 = sampleII(mm2, r11, 2)

            b1 = random.sample(mm3, 1)
            b1 = ''.join(re.findall(r'[A-Za-z]', str(b1)))
            c11.append(b1)
            c11 = sampleII(mm3, c11, 1)
            S6M = sampleM(r11, c11, mm2, mm3, RM, CM, sum1M)
            S6D = sampleD(r11, c11, mm2, mm3, RD, CD, sum1D)
            S6R = sampleR(r11, c11, mm2, mm3, RR, CR, sum1R)
            S6C = sampleC(r11, c11, mm2, mm3, RC, CC, sum1C)
            MSE66[ii] = mean_squared_error(RS, [S6M, S6D, S6R, S6C])
            # Spearman
            Y1 = pd.Series([S6M, S6D, S6R, S6C])
            Sr66[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            r11.extend(c11)
            E6[ii] = findee(r11) + 5

            # # select 7
            r11 = []  # 7
            c11 = []  # 14
            a1 = random.sample(mm2, 1)
            a1 = ''.join(re.findall(r'[A-Za-z]', str(a1)))
            r11.append(a1)
            r11 = sampleII(mm2, r11, 4)

            b1 = random.sample(mm3, 1)
            b1 = ''.join(re.findall(r'[A-Za-z]', str(b1)))
            c11.append(b1)
            c11 = sampleII(mm3, c11, 3)
            S7M = sampleM(r11, c11, mm2, mm3, RM, CM, sum1M)
            S7D = sampleD(r11, c11, mm2, mm3, RD, CD, sum1D)
            S7R = sampleR(r11, c11, mm2, mm3, RR, CR, sum1R)
            S7C = sampleC(r11, c11, mm2, mm3, RC, CC, sum1C)
            MSE77[ii] = mean_squared_error(RS, [S7M, S7D, S7R, S7C])
            # Spearman
            Y1 = pd.Series([S7M, S7D, S7R, S7C])
            Sr77[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            r11.extend(c11)
            E7[ii] = findee(r11) + 5

            # # select 8
            r11 = []  # 8
            c11 = []  # 16
            a1 = random.sample(mm2, 1)
            a1 = ''.join(re.findall(r'[A-Za-z]', str(a1)))
            r11.append(a1)
            r11 = sampleII(mm2, r11, 5)

            b1 = random.sample(mm3, 1)
            b1 = ''.join(re.findall(r'[A-Za-z]', str(b1)))
            c11.append(b1)
            c11 = sampleII(mm3, c11, 7)
            S8M = sampleM(r11, c11, mm2, mm3, RM, CM, sum1M)
            S8D = sampleD(r11, c11, mm2, mm3, RD, CD, sum1D)
            S8R = sampleR(r11, c11, mm2, mm3, RR, CR, sum1R)
            S8C = sampleC(r11, c11, mm2, mm3, RC, CC, sum1C)
            MSE88[ii] = mean_squared_error(RS, [S8M, S8D, S8R, S8C])
            # Spearman
            Y1 = pd.Series([S8M, S8D, S8R, S8C])
            Sr88[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            r11.extend(c11)
            E8[ii] = findee(r11) + 4

            # # select 9
            r11 = []  # 9
            c11 = []  # 18
            a1 = random.sample(mm2, 1)
            a1 = ''.join(re.findall(r'[A-Za-z]', str(a1)))
            r11.append(a1)
            r11 = sampleII(mm2, r11, 8)

            b1 = random.sample(mm3, 1)
            b1 = ''.join(re.findall(r'[A-Za-z]', str(b1)))
            c11.append(b1)
            c11 = sampleII(mm3, c11, 10)
            S9M = sampleM(r11, c11, mm2, mm3, RM, CM, sum1M)
            S9D = sampleD(r11, c11, mm2, mm3, RD, CD, sum1D)
            S9R = sampleR(r11, c11, mm2, mm3, RR, CR, sum1R)
            S9C = sampleC(r11, c11, mm2, mm3, RC, CC, sum1C)
            MSE99[ii] = mean_squared_error(RS, [S9M, S9D, S9R, S9C])

            # Spearman
            Y1 = pd.Series([S9M, S9D, S9R, S9C])
            Sr99[ii] = X1.corr(Y1, method="spearman")
            # 实验数
            r11.extend(c11)
            E9[ii] = findee(r11) + 5

            # # select 10
            r11 = []  # 10
            c11 = []  # 20
            a1 = random.sample(mm2, 1)
            a1 = ''.join(re.findall(r'[A-Za-z]', str(a1)))
            r11.append(a1)
            r11 = sampleII(mm2, r11, 9)

            b1 = random.sample(mm3, 1)
            b1 = ''.join(re.findall(r'[A-Za-z]', str(b1)))
            c11.append(b1)
            c11 = sampleII(mm3, c11, 19)
            S10M = sampleM(r11, c11, mm2, mm3, RM, CM, sum1M)
            S10D = sampleD(r11, c11, mm2, mm3, RD, CD, sum1D)
            S10R = sampleR(r11, c11, mm2, mm3, RR, CR, sum1R)
            S10C = sampleC(r11, c11, mm2, mm3, RC, CC, sum1C)
            MSE10[ii] = mean_squared_error(RS, [S10M, S10D, S10R, S10C])

            # Spearman
            Y1 = pd.Series([S10M, S10D, S10R, S10C])
            Sr10[ii] = X1.corr(Y1, method="spearman")

            # 实验数
            r11.extend(c11)

            E10[ii] = findee(r11) + 5
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
    workbook.save('Excel_testI1.xlsx')
    print("done")





















if __name__ == '__main__':
    main()