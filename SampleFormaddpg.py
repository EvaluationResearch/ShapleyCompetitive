#!/usr/bin/env python

import xlrd
import xlwt
import re
import random
import numpy as np
import pandas as pd
from sklearn.metrics import mean_squared_error

a = ["M", "D", "R"];
# b=[2,2,1,0]#1
# b=[2,2,0,0]#2
b=[10,5,1,20]#3
def randoms(x):
    test = ["","M", "R", "D", "MM", "MD",  "DD",  "MR", "DR", "RR",
            "MMM", "MMD", "MMR",  "MDD", "MDR",  "MRR",
            "DDD", "DDR", "DRR", "RRR","MMMM","MMMD","MMMR","MMDD","MMDR","MMRR",
            "MDDD","MDDR", "MDRR", "MRRR","DDDD","DDDR", "DDRR","DRRR","RRRR"]


    F = random.sample(test, x)
    P = np.random.uniform(low=0, high=1, size=x)
    return F, P








def find(index, mm3, R):
    h = 0
    for tt in range(0, len(mm3)):
        h = h + 1
        index = ''.join(re.findall(r'[A-Za-z]', str(index)))

        if index == mm3[tt]:
            break

    r1 = R[h - 1]
    return r1




def findee(r11):
    VSM = []
    VSD = []
    VSR = []


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


    eeM = list(set(VSM))
    eeR = list(set(VSR))
    eeD = list(set(VSD))

    tt = eeM + eeR + eeD
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
    print(eet)
    return len(eet)


def main():
    VS = ["","M", "R", "D", "MM", "MD",  "DD",  "MR", "DR", "RR",
            "MMM", "MMD", "MMR",  "MDD", "MDR",  "MRR","DDD", "DDR", "DRR", "RRR",
            "MMMM","MMMD","MMMR","MMDD","MMDR","MMRR","MDDD","MDDR", "MDRR", "MRRD",
            "MRRR","DDDR", "DDRR","DRRR","RRRR"]
    r0=[""]
    r1=["M", "R", "D"]
    r2=["MM", "MD",  "DD", "MR", "DR", "RR"]
    r3=["MMM", "MMD", "MMR",  "MDD", "MDR",  "MRR", "DDD", "DDR", "DRR", "RRR"]
    # r4=["MMMM","MMMD","MMMR","MMDD","MMDR","MMRR","MDDD","MDDR",
    #     "MDRR", "MRRD","DDDR", "DDRR","DRRR","RRRR"]
    r00=random.sample(r0,1)
    r11=random.sample(r1,1)
    r22=random.sample(r2,1)
    r33=random.sample(r3,1)
    # r44 = random.sample(r4, 1)
    VSS=r00+r11+r22+r33
    n = findee(VSS)
    print(n)
     # 'DM', 'RM', 'MM',]

    # n=14





if __name__ == '__main__':
    main()
