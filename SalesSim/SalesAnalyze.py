import numpy as np
import matplotlib.pyplot as plt
import itertools

import xlwings as xl
from pymsgbox import alert, confirm, prompt

from scipy.optimize import minimize, LinearConstraint, Bounds
from gekko import GEKKO

WB0 = xl.Book('d:\\GypProject2023\\SalesSim\\SalesAnalyze.xlsm')
SH_BaseInfo = WB0.sheets['BasicInfo']

SelectionSheet = WB0.selection

def 판매비율(Viet_SR=SH_BaseInfo.range('b4').value, CapaRatio = float(SH_BaseInfo.range('p17').value)):    

    try:        
        for i in range(2):
            SH_BaseInfo.range('b4').value = Viet_SR
            Sales_R0 = list(SH_BaseInfo.range('b5:b14').value)
            Productivity = list(SH_BaseInfo.range('k5:k14').value)
            Price = list(SH_BaseInfo.range('n5:n14').value)
            VC = list(SH_BaseInfo.range('l5:l14').value)
            SC = list(SH_BaseInfo.range('m5:m14').value)
            FV = SH_BaseInfo.range('b18').value

            lb = SH_BaseInfo.range('d5:d14').value
            ub = SH_BaseInfo.range('e5:e14').value

            m = GEKKO(remote=False)

            Sales_R = m.Array(m.Var, 10, value=10, lb=0, ub=30)
            for i, v in enumerate(lb):Sales_R[i].lower = v
            for i, v in enumerate(ub):Sales_R[i].upper = v

            g_Pdty = m.Array(m.Param, 10, value=1)
            for i, v in enumerate(Productivity):g_Pdty[i].value = v

            g_Price = m.Array(m.Param, 10, value=1)
            for i, v in enumerate(Price):g_Price[i].value = v

            g_VC = m.Array(m.Param, 10, value=1)
            for i, v in enumerate(VC):g_VC[i].value = v

            g_SC = m.Array(m.Param, 10, value=1)
            for i, v in enumerate(SC):g_SC[i].value = v

            m.Equation(m.sum([Ri for Ri in Sales_R]) == CapaRatio)
            m.Equation(m.sum([Ri for Ri in Sales_R[:6]]) == SH_BaseInfo.range('b4').value)
            # m.Equation(m.sum([Ri for Ri in Sales_R[7:]]) == 50)
            m.Maximize(np.sum([Ri/100 * g_Pdty[i] * (g_Price[i] - g_VC[i] - g_SC[i]) - FV for i, Ri in enumerate(Sales_R)]))

            # Solver options
            # m.options.IMODE = 6
            # Solve
            res = m.solve(disp=False)

            x = [i[0] for i in Sales_R]
            print(x)
            
            SH_BaseInfo.range('b5:b10').value = np.array(x[:6]).reshape(-1, 1)
            SH_BaseInfo.range('b12:b14').value = np.array(x[7:]).reshape(-1, 1)
            
            # Uptime Ratio Update -----
            SH_BaseInfo.range('b16').value = 90 - SH_BaseInfo.range('g20').value
                        
    except Exception as e:
            alert(e)

    return SH_BaseInfo.range('o16').value, SH_BaseInfo.range('p15').value 


if SH_BaseInfo.range('f1').value=="Capa 100그래프":    
    SH_BaseInfo.range('a1').color = (255, 0, 0)
    mY = []
    
    x = [i for i in np.arange(0.,101.,10.)]    
    for r in x: mY.append(판매비율(Viet_SR=r))
    
    
    margin = np.array(mY)[:,0]/1000000    
    YProd = np.array(mY)[:,1]/1000    
    print('mY: ', mY, margin, YProd)
    # margin = [판매비율(r)[0]/1000000 for r in x]
    # YearlyProduct = [판매비율(r)[1]/1000 for r in x]
    # margin = [r for r in x]
    
    SH_BaseInfo.range('a1').color = (255, 255, 255)
    
    fig, ax1 = plt.subplots()
    ax1.plot(x,margin, 'r', label='Margin')
    ax1.set_xlabel('Viet Sales Ratio(%)')
    ax1.set_ylabel('Total Yearly Margin(mil. VND)')
    ax1.legend(loc='upper left')
    
    ax2 = ax1.twinx()
    ax2.plot(x, YProd, 'b', label = 'Yearly Product')
    ax2.set_ylabel('Yearly Product(k sqm)')
    ax2.legend(loc='upper right')
    
    plt.grid(True)
    plt.show()
    
elif SH_BaseInfo.range('f1').value=="조업비":
    SH_BaseInfo.range('a1').color = (255, 0, 0)
    
    OPR = [i for i in np.arange(30,101, 10)]  
    
    fig, ax1 = plt.subplots()
    
    for idx, opr in enumerate(OPR):
        mY = []
        SH_BaseInfo.range('p17').value = opr
        x = [i for i in np.arange(0.,opr+1,10.)]  
        
        for r in x:     
            mY.append(판매비율(Viet_SR=r, CapaRatio=opr))
    
        margin = np.array(mY)[:,0]/1000000    
        YProd = np.array(mY)[:,1]/1000    
        print('mY: ', mY, margin, YProd)
        # margin = [판매비율(r)[0]/1000000 for r in x]
        # YearlyProduct = [판매비율(r)[1]/1000 for r in x]
        # margin = [r for r in x]
        
        SH_BaseInfo.range('a1').color = (255, 255, 255)
        
        print('idx: ', idx)
        ax1 = plt.subplot(5, 2, idx+1)
        ax1.set_title("Yearly Operation Ratio = " + str(opr))
        
        print('x:', x, ' \n margin:', margin)
        
        ax1.plot(x,margin, 'r', label='Margin')
        ax1.set_xlabel('Viet Sales Ratio(%)')
        ax1.set_ylabel('Yearly Margin(mil. VND)')
        ax1.legend(loc='upper left')
        
        ax2 = ax1.twinx()
        ax2.set_ylabel('Yearly Product(k sqm)')
        ax2.legend(loc='upper right')
        ax2.plot(x, YProd, 'b', label = 'Yearly Product')
    
    
    plt.grid(True)
    plt.show()
    
else: 
    SH_BaseInfo.range('a1').color = (255, 0, 0)
    판매비율()
    SH_BaseInfo.range('a1').color = (255, 255, 255)



