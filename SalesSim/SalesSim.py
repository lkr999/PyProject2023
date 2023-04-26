import numpy as np
import matplotlib.pyplot as plt

# import openpyxl as xl
import xlwings as xl
from pymsgbox import alert

WB0 = xl.Book('판매가시나리오.xlsx')

SH_BaseInfo = WB0.sheets['BasicInfo']

Profit = []
COL = str(SH_BaseInfo.range('b3').value)
SH_Sim = WB0.sheets[SH_BaseInfo.range(COL + '4').value]

Sales_year_sqm = np.array(
    [s for s in np.arange(2000000., 31000000., 1000000.)])
Price_sqm = [p for p in np.arange(15000., 51000., 1000.)]

Cost_repair = SH_BaseInfo.range('b30').value * 1000000

vCost_manufacture = float(SH_BaseInfo.range(COL + '5').value)
vCost_repair = Cost_repair * 0.65 / 30000000   # 변동성 고정비
vCost_shipping = float(SH_BaseInfo.range(COL + '12').value)   # 물류비
vCost_rebate = float(SH_BaseInfo.range(COL + '13').value)   # rebate

mfCost_1 = SH_BaseInfo.range(COL + '20').value * 1000000.
sfCost_1 = SH_BaseInfo.range(COL + '25').value * 1000000.

vCost_year = (vCost_manufacture + vCost_repair +
              vCost_shipping + vCost_rebate) * Sales_year_sqm
fCost_year = mfCost_1 + sfCost_1

Cost_year = vCost_year + fCost_year
if Sales_year_sqm.all() > 0:
    Cost_unit = Cost_year/Sales_year_sqm

try:
    for p in Price_sqm:
        profit = p * Sales_year_sqm - Cost_year
        Profit.append(profit)

    SH_Sim.range('a4:az1000').clear_contents()
    SH_Sim.range('h3:az3').clear_contents()

    SH_Sim.range('a4').value = Sales_year_sqm.reshape(-1, 1)/1000  # 판매량
    SH_Sim.range('b4').value = vCost_year.reshape(-1, 1)/1000000   # 변동비
    SH_Sim.range('c4').value = (
        vCost_year/Sales_year_sqm).reshape(-1, 1)   # 변동원가
    SH_Sim.range('d4').value = (
        fCost_year * np.array([1 for i in Sales_year_sqm])).reshape(-1, 1) / 1000000  # 고정비
    SH_Sim.range('e4').value = (
        fCost_year / Sales_year_sqm).reshape(-1, 1)  # 고정비
    SH_Sim.range('f4').value = Cost_year.reshape(-1, 1) / 1000000
    SH_Sim.range('g4').value = Cost_unit.reshape(-1, 1)

    SH_Sim.range('i3').value = Price_sqm
    SH_Sim.range('i4').value = np.array(Profit).T / 1000000.

    Range1 = SH_Sim.range('i4:az32').expand('down')

    if SH_BaseInfo.range('f1').value:
        for row in Range1.rows:
            for cell in row.columns:
                if cell.value is not None:
                    if cell.value > ((-1)*fCost_year/1000000) and cell.value <= 0:
                        cell.color = (255, 255, 204)
                    elif cell.value > 0:
                        cell.color = (204, 255, 204)
                    else:
                        cell.color = (255, 255, 255)

    price = [p for p in np.arange(15000., 51000., 2000.)]
    ax = plt.subplot(111)

    for p in price:
        profit = (p * Sales_year_sqm - Cost_year)
        ax.plot(Sales_year_sqm/1000, profit/1000000, lw=2, label=str(p))

        ax.set_xticks([x for x in range(0, 31000, 1000)])

        ax.set_ylabel('Profits (mil. VND/year)')
        ax.set_xlabel('Sales (k sqm/year)')

        ax.legend(loc='upper right', frameon=False)

        # mark1 = ax.annotate('Price(VND/sqm): %7.0f' % p,xy=p, xycoords='data',
        #                     xytext=p, textcoords='data', color='red', fontsize=12,
        #                     arrowprops=dict(arrowstyle="->", connectionstyle="arc3, rad = 0.3", color='red'), )

    # plt.show(True)

except Exception as e:
    alert(e)

plt.legend = True
plt.grid(True)
plt.show()

alert("Done")

# ---- Graph ---------------------
# fig = plt.figure(figsize=(9, 6))
# ax = fig.add_subplot(111, projection='3d')
#
# sale_m, price_m = np.meshgrid(Sales_year_sqm, Price_sqm)
# Profit = (price_m*sale_m - ((vCost_1 + vCost_2) * sale_m + fCost_year)) / 1000000.
# ax.plot_surface(sale_m, price_m, Profit, cmap="brg_r")
# plt.show()

WB0.save('판매가시나리오.xlsx')
