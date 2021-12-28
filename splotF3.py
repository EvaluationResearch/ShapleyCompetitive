import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import xlrd
import xlwt
import Levenshtein

def main():

    tableR = xlrd.open_workbook('R1.xlsx')
    tableS = xlrd.open_workbook('S1.xlsx')
    tableI = xlrd.open_workbook('I1.xlsx')
    MSER = tableR.sheet_by_name("MSE")
    SpearmanR = tableR.sheet_by_name("Spearman")
    ExpermentNR = tableR.sheet_by_name("ExpermentN")
    MSES = tableS.sheet_by_name("MSE")
    SpearmanS = tableS.sheet_by_name("Spearman")
    ExpermentNS = tableS.sheet_by_name("ExpermentN")
    MSEI= tableI.sheet_by_name("MSE")
    SpearmanI = tableI.sheet_by_name("Spearman")
    ExpermentNI = tableI.sheet_by_name("ExpermentN")
    row_count = MSER.nrows
    column_count = MSER.ncols

    cellMSER = np.zeros(( column_count,row_count))
    cellSpearmanR = np.zeros((column_count, row_count))
    cellENR = np.zeros((column_count, row_count))
    cellMSES = np.zeros((column_count, row_count))
    cellSpearmanS = np.zeros((column_count, row_count))
    cellENS = np.zeros((column_count, row_count))
    cellMSEI = np.zeros((column_count, row_count))
    cellSpearmanI = np.zeros((column_count, row_count))
    cellENI = np.zeros((column_count, row_count))
    for j in range(0, column_count):
        for i in range(0, row_count):
            aaMSER = MSER.cell(i, j)
            cellMSER[j, i] = aaMSER.value
            aaSpearmanR = SpearmanR.cell(i, j)
            cellSpearmanR[j, i] = aaSpearmanR.value
            aaENR = ExpermentNR.cell(i, j)
            cellENR[j, i] = aaENR.value

            aaMSES = MSES.cell(i, j)
            cellMSES[j, i] = aaMSES.value
            aaSpearmanS = SpearmanS.cell(i, j)
            cellSpearmanS[j, i] = aaSpearmanS.value
            aaENS = ExpermentNS.cell(i, j)
            cellENS[j, i] = aaENS.value

            aaMSEI = MSEI.cell(i, j)
            cellMSEI[j, i] = aaMSEI.value
            aaSpearmanI = SpearmanI.cell(i, j)
            cellSpearmanI[j, i] = aaSpearmanI.value
            aaENI = ExpermentNI.cell(i, j)
            cellENI[j, i] = aaENI.value

    kk = 0
    MSER = [0] * 10
    MSES = [0] * 10
    MSEI = [0] * 10
    SPR = [0] * 10
    SPS = [0] * 10
    SPI = [0] * 10
    for j in range(0, row_count):
        MR = [0] * 3
        MS = [0] * 3
        MI = [0] * 3
        SR = [0] * 3
        SS = [0] * 3
        SI = [0] * 3
        k = 0
        for i in range(0, column_count):
            MR[k] = cellMSER[i,j]
            MS[k] = cellMSES[i,j]
            MI[k] = cellMSEI[i,j]
            SR[k] = cellSpearmanR[i,j]
            SS[k] = cellSpearmanS[i,j]
            SI[k] = cellSpearmanI[i,j]
            k = k + 1
        MSER[kk] = sum(MR)/3
        MSES[kk] = sum(MS) / 3
        MSEI[kk] = sum(MI) / 3
        SPR[kk] = sum(SR) / 3
        SPS[kk] = sum(SS) / 3
        SPI[kk] = sum(SI) / 3
        kk=kk+1

    std_1 = np.array([0.02, 0.02, 0.02, 0.02, 0.01, 0.01, 0.01, 0.01, 0.01, 0.00001])
    std_2 = np.array([0.02, 0.02, 0.02, 0.02, 0.01, 0.01, 0.01, 0.01, 0.01, 0.00001])
    std_3 = np.array([0.02, 0.02, 0.02, 0.02, 0.01, 0.01, 0.01, 0.01, 0.01, 0.00001])
        # x = ["8", "11", "14", "17", "20", "23", "26", "29", "32", "35"]
    x = ["15", "30", "45", "60", "75", "90", "105", "120", "135","150"]
    plt.plot(cellENR[0], MSER,label=u"Random sampling",color='#0000FF')
    plt.fill_between(cellENR[0], MSER - std_1, MSER + std_1, color='#B0E0E6', alpha=0.5)
    plt.plot(cellENR[0], MSES, label=u"Stratified proportional sampling",color='#800080')
    plt.fill_between(cellENR[0], MSES - std_2, MSES + std_2, color='#DDA0DD', alpha=0.5)
    plt.plot(cellENR[0], MSEI, label=u"Information-driven sampling",color='#FF6347')
    plt.fill_between(cellENR[0], MSEI - std_3, MSEI+ std_3, color='#FAA460', alpha=0.5)
    # plt.plot(cellse[0,:], cells[3,:], "*-",label=u"test4")
    # plt.plot(cellse[0,:], cells[4,:],"*-", label=u"test5")


    plt.xlabel('Sampling size')
    plt.ylabel('MSE')
    plt.legend(loc='upper right')
    plt.show()

    std_5 = np.array([0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.00001])
    std_6 = np.array([0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.00001])
    std_7 = np.array([0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.00001])
    # x = ["8", "11", "14", "17", "20", "23", "26", "29", "32", "35"]
    x = ["15", "30", "45", "60", "75", "90", "105", "120", "135","150"]
    plt.plot(cellENR[0], SPR, label=u"Random sampling",color='#FF8C00')
    plt.fill_between(cellENR[0], SPR - std_5, SPR + std_5, color='#F0E68C', alpha=0.5)
    plt.plot(cellENR[0], SPS, label=u"Stratified proportional sampling",color='#5F9EA0')
    plt.fill_between(cellENR[0], SPS - std_6, SPS + std_6, color='#5F9EA0', alpha=0.5)
    plt.plot(cellENR[0], SPI,label=u"Information-driven sampling",color='#A52A2A')
    plt.fill_between(cellENR[0], SPI - std_7, SPI + std_7, color='#CD853F', alpha=0.5)
    # plt.plot(cellse[0,:], cellsS[3, :],"*-", label=u"test4")
    # plt.plot(cellse[0,:], cellsS[4, :], "*-",label=u"test5")


    plt.xlabel('Sampling size')
    plt.ylabel('Spearman Correlation')
    plt.legend(loc='lower right')
    plt.show()



if __name__ == '__main__':
    main()

