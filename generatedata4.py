from itertools import combinations
import math
import bisect
import sys
import xlrd
import xlwt
import re
import numpy as np

P=np.random.uniform(low=0.0, high=1.0, size=5)
# P=[0.74421029, 0.30133586, 0.72358195,  0.547016, 0.43886731] #test1
# P=[0.81340156, 0.62245651, 0.05931,  0.46608105, 0.98067221]#test2
# P=[0.84502405, 0.36937625, 0.02653019,  0.5268023, 0.64680392]#test3
P=[0.46712647,0.67700473,0.91128749,0.20564252]
print(P)


def findvalue():
    table = xlrd.open_workbook("16or.xlsx")
    sheet = table.sheet_by_name("test3")
    row_count = sheet.nrows
    column_count = sheet.ncols
    R=np.zeros(251)
    for j in range(0, column_count):
        for i in range(1, row_count):
            cell = sheet.cell(i, j)
            cell = str(cell).split(":")
            want = ''.join(re.findall(r'[A-Za-z]', str(cell[1])))
            T=np.zeros((len(want)))
            for k in range(len(want)):
                if want[k]=="M":
                   T[k]=P[0]
                if want[k] == "D":
                    T[k] = P[1]
                if want[k] == "R":
                    T[k] = P[2]
                if want[k] == "C":
                    T[k] = P[3]

            R[i]=sum(T)
        return R

def findvaluefT():
    table = xlrd.open_workbook("test.xlsx")
    sheet = table.sheet_by_name("Sheet4")
    row_count = sheet.nrows
    column_count = sheet.ncols
    R=np.zeros(36)
    for i in range(1, row_count):
        cell = sheet.cell(i, 0)
        cell = str(cell).split(":")
        want = ''.join(re.findall(r'[A-Za-z]', str(cell[1])))

        T = np.zeros((len(want)))
        for k in range(len(want)):
            if want[k] == "M":
                T[k] = P[0]
            if want[k] == "D":
                T[k] = P[1]
            if want[k] == "R":
                T[k] = P[2]
            if want[k] == "C":
                T[k] = P[3]

        R[i] = sum(T)
    return R




def wirteinex(R):
    workbook = xlwt.Workbook(encoding='utf-8')
    # 创建一个worksheet
    worksheet = workbook.add_sheet('My Worksheet', cell_overwrite_ok=True)

    # 写入excel
    # 参数对应 行, 列, 值
    for u in range(1, 36):
    # for u in range(1, 70):
        worksheet.write(u, 0, R[u])




    # 保存
    workbook.save('Excel_test.xls')
def main():
    # R=findvalue()
    R=findvaluefT()
    wirteinex(R)
if __name__ == '__main__':
    main()


