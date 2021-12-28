import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import xlrd
import xlwt


def main():

    table = xlrd.open_workbook('I51.xlsx')
    MSE = table.sheet_by_name("MSE")
    Spearman = table.sheet_by_name("Spearman")
    ExpermentN = table.sheet_by_name("ExpermentN")
    row_count = MSE.nrows
    column_count = MSE.ncols
    column_countS = Spearman.ncols
    row_countS = Spearman.nrows

    cells = np.zeros(( column_count,row_count))
    cellsS = np.zeros((column_countS, row_countS))
    cellse = np.zeros((column_countS, row_countS))
    for j in range(0, column_count):
        for i in range(0, row_count):
            aa = MSE.cell(i, j)
            cells[j, i] = aa.value
            aaS = Spearman.cell(i, j)
            cellsS[j, i] = aaS.value
            aae = ExpermentN.cell(i, j)
            cellse[j, i] = aae.value

    std_1 = np.array([0.02, 0.02, 0.02, 0.02, 0.01,0.01,  0.01,0.01,0.01, 0.00001])
    std_2 = np.array([0.02, 0.02, 0.02, 0.02, 0.01, 0.01,  0.01,0.01, 0.01, 0.00001])
    std_3 = np.array([0.02, 0.02, 0.02, 0.02, 0.01, 0.01,  0.01,0.01, 0.01, 0.00001])
    # x = ["8", "11", "14", "17", "20", "23", "26", "29", "32", "35"]
    x = ["15", "30", "45", "60", "75", "90", "105", "120", "135","150"]
    plt.plot(cellse[0,:],cells[0,:], label=u"test1", alpha=0.5, color='#0000FF')
    plt.fill_between(cellse[0,:], cells[0,:] - std_1, cells[0,:] + std_1, color='#B0E0E6', alpha=0.5)
    plt.plot(cellse[0,:], cells[1,:], label=u"test2", alpha=0.5,color='#800080')
    plt.fill_between(cellse[0,:], cells[1,:] - std_2, cells[1,:] + std_2, color='#DDA0DD', alpha=0.5)
    plt.plot(cellse[0,:], cells[2,:], label=u"test3", alpha=0.5,color='#FF6347')
    plt.fill_between(cellse[0, :], cells[2, :] - std_3, cells[2, :] + std_3, color='#FAA460', alpha=0.5)
    # plt.plot(cellse[0,:], cells[3,:], "*-",label=u"test4")
    # plt.plot(cellse[0,:], cells[4,:],"*-", label=u"test5")


    plt.xlabel('Sampling size')
    plt.ylabel('MSE')
    plt.legend(loc='upper right')
    plt.show()
    # x = ["8", "11", "14", "17", "20", "23", "26", "29", "32", "35"]
    x = ["15", "30", "45", "60", "75", "90", "105", "120", "135","150"]

    std_5= np.array([0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.00001])
    std_6 = np.array([0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.00001])
    std_7 = np.array([0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01, 0.00001])
    plt.plot(cellse[0,:], cellsS[0, :], label=u"test1",color='#FF8C00')
    plt.fill_between(cellse[0, :], cellsS[0, :] - std_5, cellsS[0, :] + std_5, color='#F0E68C', alpha=0.5)
    plt.plot(cellse[0,:], cellsS[1, :], label=u"test2",color='#5F9EA0')
    plt.fill_between(cellse[0, :], cellsS[1, :] - std_6, cellsS[1, :] + std_6, color='#5F9EA0', alpha=0.5)
    plt.plot(cellse[0,:], cellsS[2, :], label=u"test3",color='#A52A2A')
    plt.fill_between(cellse[0, :], cellsS[2, :] - std_7, cellsS[2, :] + std_7, color='#CD853F', alpha=0.5)
    # plt.plot(cellse[0,:], cellsS[3, :],"*-", label=u"test4")
    # plt.plot(cellse[0,:], cellsS[4, :], "*-",label=u"test5")


    plt.xlabel('Sampling size')
    plt.ylabel('Spearman Correlation')
    plt.legend(loc='lower right')
    plt.show()



if __name__ == '__main__':
    main()

