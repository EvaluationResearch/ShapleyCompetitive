import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import xlrd
import xlwt
import Levenshtein
import matplotlib.pyplot as plt

def main():
    x = ["37","55","73","91","109","127", "145", "163", "181", "199"]
    PredatorM = [911561.3,1000249.3,1280749.7,1317946.7,1410450.8,1565460.1, 1477556.6, 1443646.1, 1447902.3, 1386746.3]
    PredatorD = [1786889.9,1594389.4,1484609.3,1398009.9,1335922.6,1475526.2, 1453380.3, 1424042.9, 1489382.3, 1413921.0]
    PredatorR = [441285.4,358975.9,496417.4,547682.2,562703.6,522581.7, 535628.1, 565956.7, 528051.7, 507399.4]

    plt.plot(x, PredatorM, label=u"M Shapley for Predator", color='#0000FF')
    plt.plot(x, PredatorD, label=u"D Shapley for Predator", color='#800080')
    plt.plot(x, PredatorR, label=u"R Shapley for Predator", color='orange')


    plt.xlabel('Sampling size')
    plt.ylabel('Shapley value')
    plt.legend(loc='center right')
    plt.show()

    preyM = [132612.1,146900.9,158947.9,102818.2,49100.0,34307.6, -53607.0, -79967.4, -95500.0, -103000.0]
    preyD = [-357452.7,-224957.8,-185086.8,-222604.6,-101759.7,-193282.7, -151756.9, -119664.1, -99228.7, -79186.6]
    preyR = [-3073197.9,-3130495.1,-3426728.5,-3582548.2,-3802924.4,-3816551.7, -3727675.2, -3587752.7, -3403270.3, -3361053.4]
    totalM = ["720423.0", "720403.8", "700096.0", "749228.1", "713556.0"]
    totalD = ["606851.9", "672022.4", "662223.2", "667227.1", "631295.9"]
    totalR = ["-1477274.2", "-1465052.3", "-1340468.9", "-1313693.9", "-1335857.0"]
    plt.plot(x, preyM, label=u"M Shapley for prey", color='#0000FF')
    plt.plot(x, preyD, label=u"D Shapley for prey", color='#800080')
    plt.plot(x, preyR, label=u"R Shapley for prey", color='orange')
    plt.xlabel('Sampling size')
    plt.ylabel('Shapley value')
    plt.legend(loc='center right')
    plt.show()

    M=[0.739757,0.783949,0.844287,0.861938,0.869437,0.857753,0.852921,0.848589,0.833795,0.838314]
    D=[0.779393,0.807472,0.830016,0.837306,0.847820,0.828254,0.841170,0.842696,0.837653,0.843653]
    R=[0.361566,0.369265,0.399397,0.414635,0.418695,0.403114,0.409521,0.412779,0.401758,0.405076]
    plt.plot(x, M, label=u"M Shapley", color='#0000FF')
    plt.plot(x, D, label=u"D Shapley", color='#800080')
    plt.plot(x, R, label=u"R Shapley", color='orange')
    plt.xlabel('Sampling size')
    plt.ylabel('Shapley value')
    plt.legend(loc='center right')
    plt.show()
if __name__ == '__main__':
    main()

