import sys
from tkinter.constants import W
import pandas as pd
import numpy as np
import datetime
import xlwings as xl
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
WB_Base = ROOT_DIR + 'Daily Report_BaseF.xlsx'

XLS = 'd:\\G_FactoryDB_Asset\\XLS\\'
WB_prodDB = xl.Book(XLS + 'ProductionDB.xlsb')

USER_1     = 'leekr2'
PASSWORD_1 = 'g1234'
HOST     = '10.50.3.163'
DB       = 'gfactoryDB'

# font_name = font_manager.FontProperties(fname="c:/Windows/fonts/malgun.ttf").get_name()
# rc('font', family=font_name)

# Ui_frmMain_zietGypsum = uic.loadUiType('frm_zeitGypsum.ui')


class Zeit_Gypsum(QMainWindow):
    def __init__(self, parent=None):
        # QMainWindow.__init__(self, parent=None)       
        
        super().__init__(parent)
        self.ui = uic.loadUi('D:\\PyProject2023\\ZeitGypsum\\frm_zeitGypsum2023.ui', self)
        self.ui.show()
        self.ui.setWindowTitle('Zeit Gypsum')
                     
        # Const -------------------------
        NOW = datetime.datetime.now()
        DATE_NOW = NOW.day
        self.edtDayFrom.setDate(NOW)
        self.edtDayTo.setDate(NOW)
        
        self.STEP = 0
        
        
    def ProdTUpdate(self): 
        SH_ProdT = WB_prodDB.sheets['Prod_T']
        
        count_xl_data = SH_ProdT.range('c4').value
        rowN = int(7 + count_xl_data - 1)

        curCHK = self.cur.execute("select * from Prod_T where KeyInCode=%s", str(SH_ProdT.range('b4').value))
        self.conn.commit()

        # DB upload-------------------------------------------------------
        try:
            # Delete -----
            if curCHK>0: 
                CHK_Replace = confirm('Data Exist. Surely Replace them?', 'Confirm Replace Data', buttons=['Yes, Sure', 'No'])
                
                if CHK_Replace=='Yes, Sure':    
                    sqltxt_del_1 = "DELETE FROM Prod_T WHERE KeyInCode=%s"
                    self.cur.execute(sqltxt_del_1, str(SH_ProdT.range('b4').value))
                    self.conn.commit()        
                    # Update Data ------------------
                    df_prod_T = pd.DataFrame(SH_ProdT.range('a7:ay' + str(rowN)).value, columns=SH_ProdT.range('a5:ay5').value)
                    df_prod_T['KeyInCode'] = SH_ProdT.range('b4').value
                    
                    df_prod_T.to_sql('Prod_T', con=self.engine, if_exists='append', index=False)
                    # conn.commit()
                    self.conn.close()
                    alert('Data Updated')
                else: pass
            else:
                print(curCHK, 'CHk OK 2 ----')
                df_prod_T = pd.DataFrame(SH_ProdT.range('a7:ay' + str(rowN)).value, columns=SH_ProdT.range('a5:ay5').value)
                df_prod_T['KeyInCode'] = SH_ProdT.range('b4').value
                
                df_prod_T.to_sql('Prod_T', con=self.engine, if_exists='append', index=False)
                self.conn.commit()
                self.conn.close()
                alert('Data Updated')
                
        except Exception as e:
            alert(e)
        
        pass
    
    def ProdTDataRead(self): pass
    
    def MenuAction(self, qaction): 
        _action = qaction.text()
        
        try:
            if self.cbServerConnected:                
                # Product Table to df ----------------------------------
                df_ProdT = pd.read_sql("select *,  \
                                        N_Knife*Width*Length/1000000 sqm_Knife, \
                                        N_DryerInput*Width*Length/1000000 sqm_DryerInput, \
                                        N_DryerReject*Width*Length/1000000 sqm_DryerReject, \
                                        N_SampleReject*Width*Length/1000000 sqm_SampleReject, \
                                        N_Stacker*Width*Length/1000000 sqm_Stacker, \
                                        (N_Knife - N_DryerInput)*Width*Length/1000000 sqm_Loss_Wetend \
                                        from Prod_T;", self.conn)
                
                cond_date = (self.edtDayFrom.date().toPyDate().strftime('%Y%m%d') <= df_ProdT['Date']) &  (self.edtDayTo.date().toPyDate().strftime('%Y%m%d') >= df_ProdT['Date'])

                print(df_ProdT)
   
        except Exception as e: alert(e)
        
            
        if _action == 'Read From DB':
            try:
                SH_ReadDB = WB_prodDB.sheets['ReadDB']
                SH_ReadDB.range('a1').select
                SH_ReadDB.range('a5:bz10000').clear_contents()
                SH_ReadDB.range('a5').value = df_ProdT 
            except Exception as e: alert(e)

        if _action == 'DailyProd':
            try:
                SH_DailyProd = WB_prodDB.sheets['DailyProd']
                SH_DailyProd.range('a1').select
                SH_DailyProd.range('a5:bz10000').clear_contents()            
                
                gbDailyProd = df_ProdT[cond_date].groupby(['Date'])['sqm_Knife','sqm_DryerInput','sqm_DryerReject','sqm_SampleReject','sqm_Stacker','sqm_Loss_Wetend','Good','NoGood','Sort','ReCut'].sum()
                    
                gbDailyProd['ratio_WetendLoss'] = gbDailyProd.apply(lambda x: x['sqm_DryerInput']/x['sqm_Knife'] ,  axis=1)
                
                SH_DailyProd.range('a5').value = gbDailyProd
                print('CHK Ok ----')
                
            except Exception as e: alert(e)
        
    
    def ServerLogin(self):
        # DB Connect -------------
        if self.edtID.text() == USER_1 and self.edtPassword.text() == PASSWORD_1:
            try: 
                self.ConnectDB(USER=USER_1, HOST=HOST, PASSWORD=PASSWORD_1, DB=DB)
                if self.conn.ping:
                    self.cbServerConnected.setChecked(True)
                    self.cbServerConnected.setStyleSheet("QCheckBox:unchecked{ color: yello; }QCheckBox:checked{ color: red; }")
                else: alert('Check the DB Connection')
            except Exception as e: alert(e)            
        else: alert('Wrong ID or Password, Please Check again!!')     
        return
    
        
    def ConnectDB(self, USER, HOST, PASSWORD, DB):
        # DB Setting ------

        try:
            # self.conn = pymysql.connect(host='192.168.1.95', user=USER, password=PASSWORD, db='zeitgypsumdb', charset='utf8')
            self.conn = pymysql.connect(host=HOST, user=USER, password=PASSWORD, db=DB, charset='utf8')
            # Get Cursor
            pymysql.install_as_MySQLdb()
            self.engine = create_engine("mysql://{user}:{password}@{host}/{db}".format(user=USER, password=PASSWORD, host=HOST, db=DB))
            # self.conn = self.engine.connect()
            self.cur = self.conn.cursor()

        except mariadb.Error as e:
            print(f"Error connecting to MariaDB Platform: {e}")
            sys.exit(1)

        


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(True)
    main_win = Zeit_Gypsum()
    main_win.show()
    sys.exit(app.exec_())
    pass