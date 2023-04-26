# -*- coding: utf-8 -*-
'''
Created on 2016. 4. 30.

@author: KwangRyeol
'''
import sys
from tkinter.constants import W
import pandas as pd
import numpy as np
import datetime
import xlwings as xl
import PyQt5
import os
import mariadb
import pymysql
import qrcode
import matplotlib.pyplot as plt

from pymsgbox import alert, confirm, password, prompt

from PyQt5 import QtGui,QtCore,uic
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QMessageBox, QMenu, QMenuBar
from PyQt5.QtCore import QDateTime, Qt
from sqlalchemy import create_engine
from xlwings.utils import rgb_to_int

from PyQt5.QtWidgets import QMainWindow, QAction, qApp, QApplication



# Const --------------------------------------
ROOT_DIR = 'D:\\ZeitGypsum\\'
XL_BACKUP = ROOT_DIR + "XL_BACKUP\\"
XL_TAKEOFF = XL_BACKUP + "XL_TAKEOFF\\"
QCODE = ROOT_DIR + 'QR_CODE\\'
XLS = 'd:\\G_FactoryDB_Asset\\XLS\\'

WB_Base = ROOT_DIR + 'Daily Report_BaseF.xlsx'
WB_prodDB = xl.Book(XLS + 'ProductionDB.xlsb')

USER_1     = 'leekr2'
PASSWORD_1 = 'g1234'
HOST     = '10.50.3.163'
DB       = 'gfactoryDB'

# font_name = font_manager.FontProperties(fname="c:/Windows/fonts/malgun.ttf").get_name()
# rc('font', family=font_name)

# Ui_frmMain_zietGypsum = uic.loadUiType('frm_zeitGypsum.ui')

from frm_zeitGypsum import Ui_frmMain_zeitGypsum

class Zeit_Gypsum(QMainWindow, Ui_frmMain_zeitGypsum):
    def __init__(self, parent=None):
        QMainWindow.__init__(self, parent=None)       
        

        self.ui = Ui_frmMain_zeitGypsum()
        self.ui.setupUi(self)
        self.setWindowTitle('Zeit Gypsum')
        
             
        # Const -------------------------
        NOW = datetime.datetime.now()
        DATE_NOW = NOW.day
        self.ui.edtDate.setDate(NOW)
        self.STEP = 0

        self.XL_TakeOff_List = os.listdir(XL_TAKEOFF)
        

        #-----Column List --------
        self.t_insp_col_list = ["Date", "ShiftType", "ShiftName", "SheetNo", "Inspector", "Remark", "codeShift"]
        self.t_results_col_list = ["TOB_No", "TOB", "Thick", "Width", "Length", "Area", "Lot_No", "Time","Accept","Piler_A",
                                   "Piler_B", "WareHouse", "Thick_1", "Thick_2", "Thick_3", "Thick_4", "Thick_5","Width_A","Width_B","Length_A",
                                   "Length_B", "Diagon_A", "Diagon_B", "Angle_L", "Angle_R", "Adhesive_Face", "Adhesive_Back","Moisture","Bending_Force_MD","Bending_Force_CD",
                                   "Weight", "Density", "Sqm_Mass", "Defect_Name", "Defect_Num", "Insp_Shift_ID"]
        
        # Inspection Pro ------------------------------------------------
        self.ui.btnServerLogin.clicked.connect(self.ServerLogin)
        # self.ui.btnSaveReport.clicked.connect(self.SaveReport)
        # self.ui.btnCard.clicked.connect(self.WarrantyCard)
        self.ui.btnProdTUpLoad.clicked.connect(self.ProdTUpLoad)
        self.ui.btnProdTRead.clicked.connect(self.ProdTRead)  
        
        
        # Action ----------------------------------------------------------

        
    def ProdTRead(self):
        
        SH_ReadDB = WB_prodDB.sheets['ReadDB']
        
        try:
            if self.ui.cbServerConnected:
                df_ProdT = pd.read_sql("select * from Prod_T;", self.conn)
                print(df_ProdT)
                SH_ReadDB.range('a4').value = df_ProdT
                
                
        except Exception as e: alert(e)
        
        
        
    def ProdTUpLoad(self):
        SH_ProdT = WB_prodDB.sheets['Prod_T']
        
        count_xl_data = SH_ProdT.range('c8').value
        rowN = int(11 + count_xl_data - 1)

        curCHK = self.cur.execute("select * from Prod_T where KeyInCode=%s", str(SH_ProdT.range('b8').value))
        self.conn.commit()

        # DB upload-------------------------------------------------------
        try:
            # Delete -----
            print(curCHK, 'CHk OK ----')
            if curCHK>0: 
                CHK_Replace = confirm('Data Exist. Surely Replace them?', 'Confirm Replace Data', buttons=['Yes, Sure', 'No'])
                
                if CHK_Replace=='Yes, Sure':    
                    sqltxt_del_1 = "DELETE FROM Prod_T WHERE KeyInCode=%s"
                    self.cur.execute(sqltxt_del_1, str(SH_ProdT.range('b8').value))
                    self.conn.commit()        
                    # Update Data ------------------
                    df_prod_T = pd.DataFrame(SH_ProdT.range('a11:ay' + str(rowN)).value, columns=SH_ProdT.range('a9:ay9').value)
                    df_prod_T['KeyInCode'] = SH_ProdT.range('b8').value
                    
                    df_prod_T.to_sql('Prod_T', con=self.engine, if_exists='append', index=False)
                    # conn.commit()
                    self.conn.close()
                    alert('Data Updated')
                else: pass
            else:
                print(curCHK, 'CHk OK 2 ----')
                df_prod_T = pd.DataFrame(SH_ProdT.range('a11:ay' + str(rowN)).value, columns=SH_ProdT.range('a9:ay9').value)
                df_prod_T['KeyInCode'] = SH_ProdT.range('b8').value
                print(df_prod_T["KeyInCode"])
                
                df_prod_T.to_sql('Prod_T', con=self.engine, if_exists='append', index=False)
                self.conn.commit()
                self.conn.close()
                alert('Data Updated')
                
        except Exception as e:
            alert(e)
        
        
    def ServerLogin(self):
        # DB Connect -------------
        if self.ui.edtID.text() == USER_1 and self.ui.edtPassword.text() == PASSWORD_1:
            try: 
                self.ConnectDB(USER=USER_1, HOST=HOST, PASSWORD=PASSWORD_1, DB=DB)
                if self.conn.ping:
                    self.ui.cbServerConnected.setChecked(True)
                    self.ui.cbServerConnected.setStyleSheet("QCheckBox:unchecked{ color: yello; }QCheckBox:checked{ color: red; }")
                else: alert('Check the DB Connection')
            except Exception as e: alert(e)            
        else: alert('Wrong ID or Password, Please Check again!!')  
        
    
    def WarrantyCard(self):
        fileName = self.ui.edtDate.text() + '_' + self.ui.cbShiftType.currentText() + '_' + self.ui.cbShiftName.currentText() + '_' + self.ui.cbSheetNo.currentText() + '_TakeOFF.xlsx'
        codeShift = 'Insp_' + self.ui.edtDate.text() + '_' + self.ui.cbShiftType.currentText() + '_' + self.ui.cbShiftName.currentText() + '_' + self.ui.cbSheetNo.currentText()

        WB0 = xl.Book(XL_TAKEOFF + fileName)
        SH_TakeOff = WB0.sheets["TakeOff"]
        SH_Card    = WB0.sheets['Card']

        # Make a Inspection Result DataFrame ---------------
        List_ShiftTable = [[SH_TakeOff.range("c3").value, SH_TakeOff.range("g3").value, SH_TakeOff.range("g4").value,
                            SH_TakeOff.range("g5").value, SH_TakeOff.range("o3").value, SH_TakeOff.range("t3").value, codeShift]]
        arr_ShiftTable = np.array(List_ShiftTable)

        df_ShiftTable = pd.DataFrame(arr_ShiftTable, columns=self.t_insp_col_list)

        # Issp_Reports table ----------------------------------
        List_1 = list(SH_TakeOff.range("a10:f40").value)
        List_2 = list(SH_TakeOff.range("g10:l40").value)
        List_3 = list(SH_TakeOff.range("m10:r40").value)

        TOB_list_1 = list(SH_TakeOff.range("a8:f8").value)
        TOB_list_2 = list(SH_TakeOff.range("g8:l8").value)
        TOB_list_3 = list(SH_TakeOff.range("m8:r8").value)

        Property_list    = list(np.array(SH_TakeOff.range("u7:ab28").value).T)
        BadnRecheck_list = list(np.array(SH_TakeOff.range("u30:ab32").value).T)

        Lot_list = List_1 + List_2 + List_3
        TOB_list = [TOB_list_1] + [TOB_list_2] + [TOB_list_3]

        inspResult_list = []
        temp_TOB      = []

        for n, TOB_L in enumerate(TOB_list):
            if n>0 and (TOB_L[0]==None or TOB_L[1]==None or TOB_L[2]==None or TOB_L[3]==None):temp_TOB = TOB_list[n-1]
            else: temp_TOB = TOB_L

            k = 1

            for m, Lot_l in enumerate(Lot_list):
                if k <= 31:
                    list_Roport_1 = temp_TOB + Lot_list[m + n*31]
                    inspResult_list.append(list_Roport_1)
                    k += 1
                else:break


        colName_1 = ["TOB_No", "TOB", "Thick", "Width", "Length", "Area", "Lot_No", "Time","Accept","Piler_A","Piler_B", "WareHouse", ]
        colName_2 = ["Lot_No",  "Thick_1", "Thick_2", "Thick_3", "Thick_4", "Thick_5","Width_A","Width_B","Length_A","Length_B", "Diagon_A", "Diagon_B", "Angle_L", "Angle_R", "Adhesive_Face", "Adhesive_Back","Moisture","Bending_Force_MD","Bending_Force_CD","Weight", "Density", "Sqm_Mass",]
        colName_3 = ["Lot_No", "Defect_Name", "Defect_Num"]

        df_tobName = pd.DataFrame(WB0.sheets['Set'].range('p1:r11').value, columns=['NameCode', 'Type', 'Shape'])

        df_insp = pd.DataFrame(np.array(inspResult_list), columns=colName_1)
        df_insp = df_insp.dropna(subset=['Piler_A', 'Piler_B'], how='all')
        df_insp[['Piler_A', 'Piler_B']] = df_insp[['Piler_A', 'Piler_B']].fillna(0)
        df_insp['codeShift'] = codeShift
        df_dailyProduct = df_ShiftTable.merge(df_insp, left_on='codeShift', right_on='codeShift', how='left')


        # Card Writing -----------------------------------
        SH_TakeOff.select()
        cells = WB0.app.selection
        colN = cells.column
        rowN = cells.row
        printList = cells.value
        rgb =  (255,255,255)

        # Cell Color initialize --------------------------
        for r in range(10, 41):
            for c in [1,7,13]:
                if SH_TakeOff.range(r,c+2).value == 'X': 
                    rgb = (255,102,102)
                    print('CHK OK')
                
                elif SH_TakeOff.range(r,c+2).value == 'RC': rgb = (255,255,153)
                else: rgb =  (255,255,255)
                SH_TakeOff.range((r,c+1),(r,c+5)).color = rgb


        if (colN in [1,7,13]) and (rowN in range(10,41)):COND_START = True
        else:
            COND_START = False
            confirm('Please check again Selected Cells ')

        if COND_START:
            if cells.count==1: printList = [int(printList)]
            for r, k in enumerate(printList):
                try:
                    lot = int(k)
                    piler = ['Piler_A', 'Piler_B']

                    for p in piler:

                        if df_dailyProduct[df_dailyProduct['Lot_No']==lot][p].values[0] > 0:
                            part1 = df_insp[df_insp['Lot_No']==lot][['TOB','Thick', 'Width','Length']]
                            part2 = df_dailyProduct[df_dailyProduct['Lot_No']==lot][['Date','ShiftType', 'ShiftName','Inspector', 'Lot_No', 'WareHouse', p]]
                            SH_Card.range('c6').value = np.array(part1)
                            SH_Card.range('b8').value = np.array(part2)

                            qTxt = np.array(df_dailyProduct[df_dailyProduct['Lot_No']==lot]['codeShift'])[0]
                            qCode = qTxt + '_lot_' + str(lot) + p[-1] 
                            qFileName =  QCODE + qCode + '.png'
                            qr = qrcode.make(qTxt)
                            qr.save(qFileName)

                            imgCell = SH_Card.range('g5:h6')

                            img = SH_Card.pictures.add(qFileName)
                            img.width = 100
                            img.height = 100
                            img.top = imgCell.top+2
                            img.left = imgCell.left + imgCell.width/2 - img.width/2
                            SH_Card.range('g5').value = qCode

                            print(r, rowN)

                            
                            if df_dailyProduct[df_dailyProduct['Lot_No']==lot]['Accept'].values[0] =='X': SH_Card.range('a3').color = (255,102,102)
                            elif df_dailyProduct[df_dailyProduct['Lot_No']==lot]['Accept'].values[0] =='RC': SH_Card.range('a3').color = (255,255,153)
                            else: SH_Card.range('a3').color = (255,255,255)

                            print(part2)
                            SH_Card.api.PrintOut(Preview=False)

                            img.delete()


                except Exception as e:
                    alert(e)


    def SaveReport(self):
        fileName = self.ui.edtDate.text() + '_' + self.ui.cbShiftType.currentText() + '_' + self.ui.cbShiftName.currentText() + '_' + self.ui.cbSheetNo.currentText() + '_TakeOFF.xlsx'
        codeShift = 'Insp_' + self.ui.edtDate.text() + '_' + self.ui.cbShiftType.currentText() + '_' + self.ui.cbShiftName.currentText() + '_' + self.ui.cbSheetNo.currentText()

        self.WB0 = xl.Book(XL_TAKEOFF + fileName)
        print(XL_TAKEOFF + fileName)
        SH_TakeOff = self.WB0.sheets["TakeOff"]


        List_ShiftTable = [[SH_TakeOff.range("c3").value, SH_TakeOff.range("g3").value, SH_TakeOff.range("g4").value,
                            SH_TakeOff.range("g5").value, SH_TakeOff.range("o3").value, SH_TakeOff.range("t3").value, codeShift]]
        arr_ShiftTable = np.array(List_ShiftTable)

        df_ShiftTable = pd.DataFrame(arr_ShiftTable, columns=self.t_insp_col_list)

        # Issp_Reports table ----------------------------------
        List_1 = list(SH_TakeOff.range("a10:f40").value)
        List_2 = list(SH_TakeOff.range("g10:l40").value)
        List_3 = list(SH_TakeOff.range("m10:r40").value)

        TOB_list_1 = list(SH_TakeOff.range("a8:f8").value)
        TOB_list_2 = list(SH_TakeOff.range("g8:l8").value)
        TOB_list_3 = list(SH_TakeOff.range("m8:r8").value)

        Property_list    = list(np.array(SH_TakeOff.range("u7:ab28").value).T)
        BadnRecheck_list = list(np.array(SH_TakeOff.range("u30:ab32").value).T)

        Lot_list = List_1 + List_2 + List_3
        TOB_list = [TOB_list_1] + [TOB_list_2] + [TOB_list_3]

        inspResult_list = []
        temp_TOB      = []

        for n, TOB_L in enumerate(TOB_list):
            if n>0 and (TOB_L[0]==None or TOB_L[1]==None or TOB_L[2]==None or TOB_L[3]==None):temp_TOB = TOB_list[n-1]
            else: temp_TOB = TOB_L

            k = 1

            for m, Lot_l in enumerate(Lot_list):
                if k <= 31:
                    list_Roport_1 = temp_TOB + Lot_list[m + n*31]
                    inspResult_list.append(list_Roport_1)
                    k += 1
                else:break


        colName_1 = ["TOB_No", "TOB", "Thick", "Width", "Length", "Area", "Lot_No", "Time","Accept","Piler_A","Piler_B", "WareHouse", ]
        colName_2 = ["Lot_No",  "Thick_1", "Thick_2", "Thick_3", "Thick_4", "Thick_5","Width_A","Width_B","Length_A","Length_B", "Diagon_A", "Diagon_B", "Angle_L", "Angle_R", "Adhesive_Face", "Adhesive_Back","Moisture","Bending_Force_MD","Bending_Force_CD","Weight", "Density", "Sqm_Mass",]
        colName_3 = ["Lot_No", "Defect_Name", "Defect_Num"]

        df_insp = pd.DataFrame(np.array(inspResult_list), columns=colName_1)
        df_insp = df_insp.dropna(subset=['Piler_A', 'Piler_B'], how='all')

        df_measure = pd.DataFrame(np.array(Property_list), columns=colName_2)
        df_measure = df_measure.dropna(subset=['Lot_No'], how='any')

        df_BadnRechk = pd.DataFrame(np.array(BadnRecheck_list), columns=colName_3)
        df_BadnRechk = df_BadnRechk.dropna(subset=['Lot_No'], how='any')

        df_inspResults = df_insp.merge(df_measure, left_on='Lot_No', right_on='Lot_No', how='left').merge(df_BadnRechk, left_on='Lot_No', right_on='Lot_No', how='left')
        df_inspResults['codeShift'] = codeShift


        # Aggregate -------------------
        df_Agg = df_insp.merge(df_BadnRechk, left_on='Lot_No', right_on='Lot_No', how='left')
        df_Agg[['Piler_A', 'Piler_B', 'Defect_Num']] = df_Agg[['Piler_A', 'Piler_B', 'Defect_Num']].fillna(0)


        def cal_dfAgg(tob, thick, width, length, accept, pilerA, pilerB, defectName, defectNum):

            tobText = tob + ' ' + str(int(thick)) + 'X' + str(int(width)) + 'X' + str(int(length))
            area = width * length / 1000000
            productN = pilerA + pilerB

            goods = bads = rechk = 0

            if accept == 'O': goods = pilerA + pilerB
            # else: goods = 0
            if accept == 'X': bads = pilerA + pilerB
            # else: bads = 0
            if accept == 'RC': rechk = pilerA + pilerB
            # else: rechk = 0

            if defectName=='Sample':
                productN += defectNum
                bads     += defectNum

            productArea = productN * area
            goodsArea   = goods * area
            badsArea    = bads* area
            rechkArea   = rechk * area

            return tobText, productN, goods, bads, rechk, productArea, goodsArea, badsArea, rechkArea

        df_Agg['TobText']  = df_Agg.apply(lambda x: cal_dfAgg(x.TOB, x.Thick, x.Width,x.Length,x.Accept,x.Piler_A,x.Piler_B,x.Defect_Name,x.Defect_Num)[0], axis=1)
        df_Agg['ProductN'] = df_Agg.apply(lambda x: cal_dfAgg(x.TOB, x.Thick, x.Width,x.Length,x.Accept,x.Piler_A,x.Piler_B,x.Defect_Name,x.Defect_Num)[1], axis=1)
        df_Agg['GoodsN']   = df_Agg.apply(lambda x: cal_dfAgg(x.TOB, x.Thick, x.Width,x.Length,x.Accept,x.Piler_A,x.Piler_B,x.Defect_Name,x.Defect_Num)[2], axis=1)
        df_Agg['BadsN']    = df_Agg.apply(lambda x: cal_dfAgg(x.TOB, x.Thick, x.Width,x.Length,x.Accept,x.Piler_A,x.Piler_B,x.Defect_Name,x.Defect_Num)[3], axis=1)
        df_Agg['RechkN']   = df_Agg.apply(lambda x: cal_dfAgg(x.TOB, x.Thick, x.Width,x.Length,x.Accept,x.Piler_A,x.Piler_B,x.Defect_Name,x.Defect_Num)[4], axis=1)

        df_Agg['ProductA']  = df_Agg.apply(lambda x: cal_dfAgg(x.TOB, x.Thick, x.Width,x.Length,x.Accept,x.Piler_A,x.Piler_B,x.Defect_Name,x.Defect_Num)[5], axis=1)
        df_Agg['GoodsA']    = df_Agg.apply(lambda x: cal_dfAgg(x.TOB, x.Thick, x.Width,x.Length,x.Accept,x.Piler_A,x.Piler_B,x.Defect_Name,x.Defect_Num)[6], axis=1)
        df_Agg['BadsA']     = df_Agg.apply(lambda x: cal_dfAgg(x.TOB, x.Thick, x.Width,x.Length,x.Accept,x.Piler_A,x.Piler_B,x.Defect_Name,x.Defect_Num)[7], axis=1)
        df_Agg['RechkA']    = df_Agg.apply(lambda x: cal_dfAgg(x.TOB, x.Thick, x.Width,x.Length,x.Accept,x.Piler_A,x.Piler_B,x.Defect_Name,x.Defect_Num)[8], axis=1)

        df_Agg_sum = df_Agg.groupby(['TOB_No', 'TobText'])['ProductN', 'GoodsN', 'BadsN', 'RechkN', 'ProductA', 'GoodsA', 'BadsA', 'RechkA', 'Defect_Num'].sum()

        # Aggregate value -----------------------
        self.WB0.sheets["temp"].range("a1:az10000").clear_contents()
        self.WB0.sheets["temp"].range("a2").value = df_Agg_sum
        SH_TakeOff.range("t36:ab39").clear_contents()
        SH_TakeOff.range("t36").value = self.WB0.sheets["temp"].range("b3:j6").value

        # CHK ------------------------
        CHK_txt = ''
        CHK_arr = np.array(df_Agg_sum)
        k = 0
        for i, aggSum in enumerate(CHK_arr):
            if aggSum[-1] == (aggSum[2]+aggSum[3]): msg = ''
            else:
                k += 1
                CHK_txt += ' ' + str(i+1)

        msg = 'Do Check Tob Number: ' + CHK_txt
        if k == 0:
            CHK_OK = True
            msg = ''
        else:
            CHK_OK = False

        SH_TakeOff.range('v29').value = msg

        # self.WB0.sheets["test"].range("a1:az10000").clear_contents()
        # self.WB0.sheets["test"].range("a2").value = df_Agg_sum
        # print("CHK OK")

        # DB upload-------------------------------------------------------
        if CHK_OK:
            try:
                # t_insp_shift -----
                sqltxt_del_1 = "DELETE FROM t_insp_shift WHERE codeShift=%s"
                self.cur.execute(sqltxt_del_1, codeShift)
                self.conn.commit()

                df_ShiftTable.to_sql('t_insp_shift', con=self.engine, if_exists='append', index=False)
                self.conn.commit()

                # t_insp_results -----
                sqltxt_del_2 = "DELETE FROM t_insp_results WHERE codeShift=%s"
                self.cur.execute(sqltxt_del_2, codeShift)
                self.conn.commit()

                df_inspResults.to_sql('t_insp_results', con=self.engine, if_exists='append', index=False)
                self.conn.commit()

                confirm('Data Upload Clear')
            except Exception as e:
                alert(e)
        else: confirm('Failed data upload!!')


    def OpenFile(self):

        if len(self.ui.cbShiftType.currentText())>0: chkShiftType = True
        else:
            chkShiftType = False
            alert("Input ShiftType")
        if len(self.ui.cbShiftName.currentText())>0: chkShiftName = True
        else:
            chkShiftName = False
            alert("Input ShiftName")
        if len(self.ui.cbSheetNo.currentText())>0: chkSheetNo = True
        else:
            chkSheetNo = False
            alert("Input SheetNo")

        fileName = self.ui.edtDate.text() + '_' + self.ui.cbShiftType.currentText() + '_' \
                   + self.ui.cbShiftName.currentText() + '_' \
                   + self.ui.cbSheetNo.currentText() + '_TakeOFF.xlsx'

        if fileName in self.XL_TakeOff_List:
            chkFileName = False
            alert("Same File Exist")
        else:  chkFileName = True

        if chkShiftName and chkShiftType and chkSheetNo and chkFileName:
            xl.Book(WB_Base).app.quit()
            self.WB0 = xl.Book(WB_Base)
            self.WB0.save(XL_TAKEOFF + fileName)

            SH_TakeOff = self.WB0.sheets["TakeOff"]

            SH_TakeOff.range("c3").value = self.ui.edtDate.text()
            SH_TakeOff.range("g3").value = self.ui.cbShiftType.currentText()
            SH_TakeOff.range("g4").value = self.ui.cbShiftName.currentText()
            SH_TakeOff.range("g5").value = self.ui.cbSheetNo.currentText()

            self.STEP += 1
            pass


    def ConnectDB(self, USER, HOST, PASSWORD, DB):
        # DB Setting ------

        try:
            # self.conn = pymysql.connect(host='192.168.1.95', user=USER, password=PASSWORD, db='zeitgypsumdb', charset='utf8')
            self.conn = pymysql.connect(host=HOST, user=USER, password=PASSWORD, db=DB, charset='utf8')

        except mariadb.Error as e:
            print(f"Error connecting to MariaDB Platform: {e}")
            sys.exit(1)

        # Get Cursor
        pymysql.install_as_MySQLdb()
        self.engine = create_engine("mysql://{user}:{password}@{host}/{db}".format(user=USER, password=PASSWORD, host=HOST, db=DB))
        self.cur = self.conn.cursor()
        


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(True)
    main_win = Zeit_Gypsum()
    main_win.show()
    sys.exit(app.exec_())
    pass