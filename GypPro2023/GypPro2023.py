import sys
import xlwings as xl
import numpy as np
import sympy
import scipy.integrate as sci
import matplotlib.pylab as plt

from pymsgbox import alert, confirm, password, prompt
from xlwings.utils import rgb_to_int
from sympy import symbols, solve

# Git 에 등록되어짐 -----

# Const --------------------------------------

# Stucco Symbols -------------------------------
Stucco, RG, DG, AG, CW, NonHydrate, Moisture = symbols('Stucco RG DG AG CW NonHydrate Moisture')
RG_NG, RG_FGD, RG_Scrap, Ratio_NG, Ratio_FGD, Ratio_Scrap = symbols('RG_NG RG_FGD RG_Scrap Ratio_NG Ratio_FGD Ratio_Scrap')
Moisture_mass_NG, Moisture_mass_FGD, Moisture_mass_Scrap = symbols('Moisture_mass_NG Moisture_mass_FGD Moisture_mass_Scrap')
CW_mass, DG_NG, DG_FGD, DG_Scrap, NonHydrate_NG, NonHydrate_FGD, NonHydrate_Scrap = symbols('CW_mass DG_NG DG_FGD DG_Scrap NonHydrate_NG NonHydrate_FGD NonHydrate_Scrap')
Factor_CGtoDG, RG_NG_Basic, RG_FGD_Basic, RG_Scrap_Basic, NG_Calcining, Evapor_Calcine \
    = symbols('Factor_CGtoDG RG_NG_Basic RG_FGD_Basic RG_Scrap_Basic NG_Calcining Evapor_Calcine')

Annual_Product, Thick, Width, Length, Length_Wet, Density,  Yield, Operation_Day, Uptime_Ratio, LineSpeed, Basic_Dry_Mass, Stucco_Basic, CF_Basic, BB_Basic, \
NG_Basic, FGD_Basic, Scrap_Basic\
    = symbols('Annual_Product Thick Width Length Length_Wet  Density  Yield	Operation_Day Uptime_Ratio LineSpeed Basic_Dry_Mass Stucco_Basic '
              'CF_Basic BB_Basic, NG_Basic FGD_Basic Scrap_Basic')
    
XLS_Local = 'd:\\G_FactoryDB_Asset\\XLS_Local\\'
WB_GypPro = xl.Book(XLS_Local + 'GypPro2022.xlsb')

# WB_GypPro = xl.Book("D:\\GypProject2023\\GypPro2023\\GypPro2022.xlsb")
SH_SetList = WB_GypPro.sheets['SetList']
SH_StuccoCode = WB_GypPro.sheets['StuccoCode']
SH_ValCHK = WB_GypPro.sheets['ValCHK']
SH_MaterialTable = WB_GypPro.sheets['MaterialTable']
SH_Wetend = WB_GypPro.sheets['Wetend']


SelectionSheet = WB_GypPro.selection
    

def StuccoCode():

    col_name = SH_StuccoCode.range('c1').value
    set_V =  np.array(SH_StuccoCode.range(col_name + '5:' + col_name + '29').value)

    stucco_v = set_V[0]

    ratio_ng_v = set_V[1]/100
    ratio_fgd_v = set_V[2]/100
    ratio_scrap_v = set_V[3]/100

    moisture_ng_v = set_V[6]/100
    moisture_fgd_v = set_V[7]/100
    moisture_scrap_v = set_V[8]/100

    purity_ng_v = set_V[11]/100
    purity_fgd_v = set_V[12]/100
    purity_scrap_v = set_V[13]/100

    cw_v = set_V[16]/100
    lhv_ng_v = set_V[17]
    calcine_heat_eff = set_V[18]


    sol = solve([Stucco - stucco_v,
        RG - (DG + Moisture),
        DG - (AG * 172 / 136 + NonHydrate),
        AG - ((1 - CW) * Stucco - NonHydrate),
        CW - cw_v,
        NonHydrate - (NonHydrate_NG + NonHydrate_FGD + NonHydrate_Scrap),
        Moisture - (Moisture_mass_NG + Moisture_mass_FGD + Moisture_mass_Scrap),

        RG_NG * (Ratio_NG + Ratio_FGD + Ratio_Scrap) - Ratio_NG * RG,
        RG_FGD * (Ratio_NG + Ratio_FGD + Ratio_Scrap) - Ratio_FGD * RG,
        RG_Scrap * (Ratio_NG + Ratio_FGD + Ratio_Scrap) - Ratio_Scrap * RG,
        Ratio_NG - ratio_ng_v,
        Ratio_FGD - ratio_fgd_v,
        Ratio_Scrap - ratio_scrap_v,

        Moisture_mass_NG - RG_NG * moisture_ng_v,
        Moisture_mass_FGD - RG_FGD * moisture_fgd_v,
        Moisture_mass_Scrap - RG_Scrap * moisture_scrap_v,

        CW_mass - (Stucco - (AG + NonHydrate)),

        DG_NG - (RG_NG - Moisture_mass_NG),
        DG_FGD - (RG_FGD - Moisture_mass_FGD),
        DG_Scrap - (RG_Scrap - Moisture_mass_Scrap),

        NonHydrate_NG - DG_NG * (1 - purity_ng_v),
        NonHydrate_FGD - DG_FGD * (1 - purity_fgd_v),
        NonHydrate_Scrap - DG_Scrap * (1 - purity_scrap_v),

        Factor_CGtoDG * Stucco - DG,
        RG_NG_Basic * Stucco - RG_NG,
        RG_FGD_Basic * Stucco - RG_FGD,
        RG_Scrap_Basic * Stucco - RG_Scrap,

        NG_Calcining - (RG - Stucco) * calcine_heat_eff / lhv_ng_v,
        Evapor_Calcine - (RG - Stucco)
        ]
    , [Stucco, RG, DG, AG, CW, NonHydrate, Moisture, RG_NG, RG_FGD, RG_Scrap, Ratio_NG, Ratio_FGD,
        Ratio_Scrap,
        Moisture_mass_NG, Moisture_mass_FGD, Moisture_mass_Scrap, CW_mass, DG_NG, DG_FGD, DG_Scrap,
        NonHydrate_NG, NonHydrate_FGD, NonHydrate_Scrap, Factor_CGtoDG,
        RG_NG_Basic, RG_FGD_Basic, RG_Scrap_Basic, NG_Calcining, Evapor_Calcine])

    # val = np.array([v for k,v in sol.items()])
    val = np.array([sol[0][i] for i in range(len(sol[0]))])

    StuccoCode_value = val
    SH_ValCHK.range('c3').value = val.reshape(-1,1)

    # XL Stucco Code Cal Result Keyin -----
    xl_StuccoCode_val = np.array(list(StuccoCode_value[24:27])
                                    + [StuccoCode_value[27]/StuccoCode_value[0]]
                                    + [StuccoCode_value[28]/StuccoCode_value[0]]
                                    )
    xl_StuccoCode_val_2 = np.array([StuccoCode_value[4]*100] + [StuccoCode_value[23]])


    SH_StuccoCode.range(col_name + '31').value = xl_StuccoCode_val.reshape(-1,1)
    SH_StuccoCode.range(col_name + '38').value = xl_StuccoCode_val_2.reshape(-1,1)

    alert('Done')

def SetList():
    col_name = SH_SetList.range('b1').value
    set_V = np.array(SH_SetList.range(col_name + '5:' + col_name + '20').value)
    SH_ValCHK.range('f3').value = set_V.reshape(-1,1)

    material_flowrate_v = np.array(SH_SetList.range(col_name + '27:' + col_name + '56').value)
    moisture_data_v = np.array(SH_MaterialTable.range('c7:d36').value)

    HeatEff_v = SH_SetList.range(col_name + '106').value

    # Pre Mixing --------------------
    premix_v = np.array(SH_SetList.range(col_name + '64:' + col_name + '82').value)
    SH_ValCHK.range('n3').value = premix_v.reshape(-1, 1)


    if premix_v[0]>0:  # STMP
        moisture_data_v[25,0] = ((1-premix_v[0]) + moisture_data_v[25,0]/100 * premix_v[0]) * 100.
    if premix_v[6]>0:  # WAX
        moisture_data_v[18,0] = ((1-premix_v[6]) + moisture_data_v[18,0]/100 * premix_v[6]) * 100.
    if premix_v[12]>0:  # Item1
        moisture_data_v[26,0] = ((1-premix_v[12]) + moisture_data_v[26,0]/100 * premix_v[12]) * 100.
    if premix_v[15]>0:  # Item2
        moisture_data_v[27,0] = ((1-premix_v[15]) + moisture_data_v[27,0]/100 * premix_v[15]) * 100.

    dry_mass_flow_v = (material_flowrate_v * moisture_data_v[:, 1]) - material_flowrate_v * ((moisture_data_v[:, 0] / 100) * moisture_data_v[:, 1])

    mass_flow_design = set_V[0]*60. * set_V[3]/1000. * set_V[2] * set_V[6]
    dg_flow = mass_flow_design - np.sum(dry_mass_flow_v)
    stucco_flow = dg_flow / set_V[12]

    moisture_total = np.sum(material_flowrate_v * ((moisture_data_v[:,0]/100) * moisture_data_v[:,1]))
    mass_flow_total = np.sum((material_flowrate_v * moisture_data_v[:,1])) + stucco_flow
    vapor_total = mass_flow_total - mass_flow_design

    gas_board = vapor_total * HeatEff_v / 9420.


    # self.SH_SetList.range('e27').value = dry_mass_flow_v.reshape(-1,1)

    SH_SetList.range(col_name + '84').value = np.sum(dry_mass_flow_v)
    SH_SetList.range(col_name + '85').value = mass_flow_design
    SH_SetList.range(col_name + '86').value = dg_flow
    SH_SetList.range(col_name + '87').value = stucco_flow
    SH_SetList.range(col_name + '22').value = stucco_flow
    SH_SetList.range(col_name + '88').value = moisture_total
    SH_SetList.range(col_name + '89').value = mass_flow_total
    SH_SetList.range(col_name + '90').value = vapor_total
    SH_SetList.range(col_name + '59').value = gas_board

    SetList_v = np.array(SH_SetList.range(col_name + '4:' + col_name + '120').value)
    SH_ValCHK.range('j3').value = SetList_v.reshape(-1, 1)
    


if __name__ == '__main__':
    import sys
    
    SH_SetList.range('a1').color = (255, 0, 0)

    selectPro = int(sys.argv[1])

    if selectPro == 1: StuccoCode()
    elif selectPro == 2: SetList()
    
    SH_SetList.range('a1').color = (255, 255, 255)