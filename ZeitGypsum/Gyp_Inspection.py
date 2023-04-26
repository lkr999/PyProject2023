# -*- coding: utf-8 -*-
'''
Created on 2016. 4. 30.

@author: KwangRyeol
'''
import sys
import pandas as pd
import pandas.io.sql as pdSQL
import numpy as np
import matplotlib.pyplot as plt
import pyodbc as pydb
import datetime
import xlwings as xl
import sqlite3
import PyQt5
import os
import itertools
import matplotlib.pyplot as plt
import seaborn as sns

from matplotlib import font_manager, rc
from pandas import ExcelWriter
from pymsgbox import alert, confirm, password, prompt
from PyQt5 import QtGui,QtCore,uic
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QMessageBox
from PyQt5.QtCore import QDateTime, Qt

sys.setrecursionlimit(10000)

DBNAME = "DB/Gyp_Inspect.db"
PICKLE_DIR = 'pickle_data/'
XL_DIR = 'c:/kccfc/'
WB = xl.Book(XL_DIR + '제품검사일보_Base.xlsx')

DB_DIR = 'DB/'
R0 = 'b9:am55'

font_name = font_manager.FontProperties(fname="c:/Windows/fonts/malgun.ttf").get_name()
rc('font', family=font_name)

Ui_MainWindow, QClassBasic_GI_Main = uic.loadUiType('GI_Form.ui')

class GI_Main(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(GI_Main, self).__init__()
        QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.setWindowTitle('석고보드 제품검사 시스템')


        NOW = datetime.datetime.now()
        DATE_NOW = NOW.day
        self.Product_Date.setDate(NOW)
        self.From_Date.setDate(NOW - datetime.timedelta(days=DATE_NOW-1))
        self.To_Date.setDate(NOW)
        self.Base_Date.setDate(NOW)

        # ComboBox Items --------------
        combo_data = WB.sheets['Set'].range('a2:i20').value

        검사원_List = np.array(combo_data)[:,0]
        검사원_List = [k for k in 검사원_List if k != None]
        self.qt_Worker.addItems(검사원_List)

        날씨_List = np.array(combo_data)[:,1]
        날씨_List = [k for k in 날씨_List if k != None]
        self.qt_Weather.addItems(날씨_List)

        교대구분_List = np.array(combo_data)[:,2]
        교대구분_List = [k for k in 교대구분_List if k != None]
        self.qt_Shift.addItems(교대구분_List)

        품종_List = np.array(combo_data)[:,3]
        품종_List = [k for k in 품종_List if k != None]
        self.qt_PJ.addItems(품종_List)

        두께_List = np.array(combo_data)[:,4]
        두께_List = [str(k) for k in 두께_List if k != None]
        self.qt_Thick.addItems(두께_List)

        나비_List = np.array(combo_data)[:,5]
        나비_List = [str(int(k)) for k in 나비_List if k != None]
        self.qt_Width.addItems(나비_List)

        길이_List = np.array(combo_data)[:,6]
        길이_List = [str(int(k)) for k in 길이_List if k != None]
        self.qt_Length1.addItems(길이_List)
        self.qt_Length2.addItems(길이_List)
        self.qt_Length3.addItems(길이_List)
        self.qt_Length4.addItems(길이_List)

        호기_List = np.array(combo_data)[:,7]
        호기_List = [k for k in 호기_List if k != None]
        self.qt_HOGI.addItems(호기_List)

        반_List = np.array(combo_data)[:,8]
        반_List = [str(int(k)) for k in 반_List if k != None]
        self.qt_BAN.addItems(반_List)

        WB.app.quit()

    def MENU_Action(self, qaction):
        _action = qaction.text()

        # xl.Book(XL_DIR + '검사현황분석_Write.xlsx').app.quit()
        self.WB2 = xl.Book(XL_DIR + '검사현황분석_Base.xlsx')
        # self.WB2.save(XL_DIR + '검사현황분석_Write.xlsx')
        conn = sqlite3.connect(DBNAME)

        sql_txt = "SELECT * From 생산DataT"
        검사_df = pd.read_sql(sql_txt, con=conn)

        검사_df['검사평'] = 검사_df.검사량 * 검사_df.환산계수
        검사_df['미선별평'] = 검사_df.미선별 * 검사_df.환산계수
        검사_df['불량평'] = 검사_df.불량계 * 검사_df.환산계수
        검사_df['샘플평'] = 검사_df.샘플 * 검사_df.환산계수
        검사_df['정품평'] = (검사_df.검사량-검사_df.미선별-검사_df.불량계-검사_df.샘플) * 검사_df.환산계수
        # 검사_df['규격구분'] = 검사_df.품종 +'_'+ str(검사_df.두께)+'*' + str(검사_df.나비) +'*'+ str(검사_df.길이)
        # print(검사_df.head())
        date_cond1 = (self.From_Date.date().toPyDate().strftime('%Y%m%d') <= 검사_df.일자)
        date_cond2 = (self.To_Date.date().toPyDate().strftime('%Y%m%d') >= 검사_df.일자)

        cond_Total = (date_cond1 & date_cond2)

        if _action =='관리도':
            sheet_Name = '관리도'
            self.WB2.sheets[sheet_Name].activate()
            호기구분 = ['2호기','3호기']

            cond_Total_2호기 = (date_cond1 & date_cond2 & (검사_df['호기']=='2호기'))
            cond_Total_3호기 = (date_cond1 & date_cond2 & (검사_df['호기']=='3호기'))

            Control_df = 검사_df[cond_Total][['일자', '호기','품종','치수_두께','BF_MD','BF_CD']]

            for k in 호기구분:
                Control_df[Control_df.호기==k].boxplot(by=['일자'])
                sns.distplot(Control_df[Control_df.호기==k].BF_MD.dropna())

                plt.title(k + '[] Boxplot by 일자')
                plt.grid(True)
                plt.show()






        if _action =='품질현황':

            sheet_Name = '일자별_물성평균'
            sheet_Name2 = '규격별_품질요약'
            self.WB2.sheets[sheet_Name].activate()

            cond_Total_2호기 = (date_cond1 & date_cond2 & (검사_df['호기']=='2호기'))
            cond_Total_3호기 = (date_cond1 & date_cond2 & (검사_df['호기']=='3호기'))

            집계량0 = pd.pivot_table(검사_df[cond_Total], index=['일자','품종','두께'], values=['치수_두께','치수_나비','대각','측각_좌','측각_우','함수율','BF_MD','BF_CD','무게','비중','MPSM','전흡수율'],
                                    margins=True, margins_name='합계', aggfunc=np.average)
            집계량2 = pd.pivot_table(검사_df[cond_Total_2호기], index=['일자','품종','두께'], values=['치수_두께','치수_나비','대각','측각_좌','측각_우','함수율','BF_MD','BF_CD','무게','비중','MPSM','전흡수율'],
                                    margins=True, margins_name='합계', aggfunc=np.average)
            집계량3 = pd.pivot_table(검사_df[cond_Total_3호기], index=['일자','품종','두께'], values=['치수_두께','치수_나비','대각','측각_좌','측각_우','함수율','BF_MD','BF_CD','무게','비중','MPSM','전흡수율'],
                                    margins=True, margins_name='합계', aggfunc=np.average)


            집계량10 = pd.pivot_table(검사_df[cond_Total], columns=['품종','두께', '호기'], values=['치수_두께','치수_나비','대각','측각_좌','측각_우','함수율','BF_MD','BF_CD','무게','비중','MPSM','전흡수율'],
                                    margins=True, margins_name='합계', aggfunc=np.average)
            집계량11 = 검사_df[cond_Total].groupby(['호기','품종','두께'])['치수_두께','치수_나비','대각','측각_좌','측각_우','함수율','BF_MD','BF_CD','무게','비중','MPSM','전흡수율']
            print(집계량11)

            iterables = [['치수_두께','치수_나비','대각','측각_좌','측각_우','함수율','BF_MD','BF_CD','무게','비중','MPSM','전흡수율'],['2호기','3호기','합계']]
            pivot_label_shift = pd.MultiIndex.from_product(iterables, names=['물성', '호기'])

            self.WB2.sheets[sheet_Name].range('a1').value = self.Base_Date.date().toPyDate()
            self.WB2.sheets[sheet_Name].range('a5:bd10000').clear_contents()
            self.WB2.sheets[sheet_Name].range('a5').value = 집계량0.reindex(columns=iterables[0])
            self.WB2.sheets[sheet_Name].range('q5').value = 집계량2.reindex(columns=iterables[0])
            self.WB2.sheets[sheet_Name].range('ag5').value = 집계량3.reindex(columns=iterables[0])

            self.WB2.sheets[sheet_Name2].range('a1').value = self.Base_Date.date().toPyDate()
            self.WB2.sheets[sheet_Name2].range('a5:bd10000').clear_contents()
            # self.WB2.sheets[sheet_Name2].range('a5').value = 검사_df[cond_Total].describe(include=['치수_두께','치수_나비','대각','측각_좌','측각_우','함수율','BF_MD','BF_CD','무게','비중','MPSM','전흡수율'])
            self.WB2.sheets[sheet_Name2].range('a5').value = 집계량11.describe()
            self.WB2.sheets[sheet_Name2].activate()


        if _action =='미선별현황':

            sheet_Name = '일자별_규격별_호기별_미선별현황'
            sheet_Name1 = '일자별_규격별_호기별_창고위치별_미선별현황'
            sheet_Name2 = '호기_일자_교대_반별_미선별'
            self.WB2.sheets[sheet_Name].activate()

            # cond_Total_2호기 = (date_cond1 & date_cond2 & (검사_df['호기']=='2호기'))
            # cond_Total_3호기 = (date_cond1 & date_cond2 & (검사_df['호기']=='3호기'))

            집계량0 = pd.pivot_table(검사_df[cond_Total], index=['일자','품종','두께','나비','길이'],columns=['호기'], values=['미선별'],
                                    margins=True, margins_name='합계', aggfunc=np.sum)
            집계량1 = pd.pivot_table(검사_df[cond_Total], index=['일자','품종','두께','나비','길이','창고위치'],columns=['호기'], values=['미선별'],
                                    margins=True, margins_name='합계', aggfunc=np.sum)

            집계량2 = pd.pivot_table(검사_df[cond_Total], index=['일자','교대','반'],columns=['호기'], values=['미선별평'],
                                    margins=True, margins_name='합계', aggfunc=np.sum)


            # iterables = [['검사평','정품평','불량평','샘플평','미선별평'], ['2호기', '3호기'], ['조간', '석간', '야간']]
            # pivot_label_shift = pd.MultiIndex.from_product(iterables, names=[None, '호기','교대'])
            self.WB2.sheets[sheet_Name].range('a1').value = self.Base_Date.date().toPyDate()
            self.WB2.sheets[sheet_Name].range('a5:bd10000').clear_contents()
            self.WB2.sheets[sheet_Name].range('a5').value = 집계량0

            self.WB2.sheets[sheet_Name1].range('a1').value = self.Base_Date.date().toPyDate()
            self.WB2.sheets[sheet_Name1].range('a5:bd10000').clear_contents()
            self.WB2.sheets[sheet_Name1].range('a5').value = 집계량1

            self.WB2.sheets[sheet_Name2].range('a1').value = self.Base_Date.date().toPyDate()
            self.WB2.sheets[sheet_Name2].range('a5:bd10000').clear_contents()
            self.WB2.sheets[sheet_Name2].range('a5').value = 집계량2


            # 집계량2.reindex(columns=pivot_label_shift).to_excel(writer, sheet_name = '일자별_호기교대별_검사현황', startrow=5)
            # 집계량1[new_index].to_excel(writer, sheet_name = '일자별__검사현황', startrow=5)


        if _action =='검사현황':

            sheet_Name = '일자별_호기별_검사현황'
            sheet_Name2 = '일자별_검사현황'
            self.WB2.sheets[sheet_Name].activate()

            cond_Total_2호기 = (date_cond1 & date_cond2 & (검사_df['호기']=='2호기'))
            cond_Total_3호기 = (date_cond1 & date_cond2 & (검사_df['호기']=='3호기'))

            집계량0 = pd.pivot_table(검사_df[cond_Total], index='일자', values=['검사평','정품평','불량평','샘플평','미선별평'],
                                    margins=True, margins_name='합계', aggfunc=np.sum)
            집계량2 = pd.pivot_table(검사_df[cond_Total_2호기], index='일자', values=['검사평','정품평','불량평','샘플평','미선별평'],
                                    margins=True, margins_name='합계', aggfunc=np.sum)
            집계량3 = pd.pivot_table(검사_df[cond_Total_3호기], index='일자', values=['검사평','정품평','불량평','샘플평','미선별평'],
                                    margins=True, margins_name='합계', aggfunc=np.sum)

            집계량4 = pd.pivot_table(검사_df[cond_Total], index='일자', columns = '호기', values=['검사평','정품평','불량평','샘플평','미선별평'],
                                    margins=True, margins_name='합계', aggfunc=np.sum)
            집계량5 = pd.pivot_table(검사_df[cond_Total], index='일자', columns = ['호기','교대'], values=['검사평','정품평','불량평','샘플평','미선별평'],
                                    margins=True, margins_name='합계', aggfunc=np.sum)


            my_order = ['검사평','정품평','불량평','샘플평','미선별평']

            # iterables = [['검사평','정품평','불량평','샘플평','미선별평'], ['2호기', '3호기'], ['조간', '석간', '야간']]
            # pivot_label_shift = pd.MultiIndex.from_product(iterables, names=[None, '호기','교대'])
            self.WB2.sheets[sheet_Name].range('a1').value = self.Base_Date.date().toPyDate()
            self.WB2.sheets[sheet_Name].range('a5:bd10000').clear_contents()
            self.WB2.sheets[sheet_Name].range('a5').value = 집계량0.reindex(columns=my_order)
            self.WB2.sheets[sheet_Name].range('h5').value = 집계량2.reindex(columns=my_order)
            self.WB2.sheets[sheet_Name].range('o5').value = 집계량3.reindex(columns=my_order)

            first_2_levels = [x.tolist() for x in 집계량4.columns.levels[:3]]
            first_2_levels[0] = my_order
            new_index = list(itertools.product(*first_2_levels))
            # 일자별품질 = 검사_df[cond_Total].groupby('일자')['BF_MD','BF_CD','측각_좌','측각_우','비중','MPSM' ].mean()

            self.WB2.sheets[sheet_Name2].range('a1').value = self.Base_Date.date().toPyDate()
            self.WB2.sheets[sheet_Name2].range('a4:p10000').clear_contents()
            self.WB2.sheets[sheet_Name2].range('a4').value = 집계량4[new_index]

            my_order2 = ['검사평','정품평','불량평','샘플평','미선별평']

            iterables = [['검사평','정품평','불량평','샘플평','미선별평'], ['2호기', '3호기'], ['조간', '석간', '야간']]
            pivot_label_shift = pd.MultiIndex.from_product(iterables, names=[None, '호기','교대'])

            self.WB2.sheets[sheet_Name2].range('u3:bd10000').clear_contents()
            self.WB2.sheets[sheet_Name2].range('u3').value = 집계량5.reindex(columns=pivot_label_shift)

            # 집계량2.reindex(columns=pivot_label_shift).to_excel(writer, sheet_name = '일자별_호기교대별_검사현황', startrow=5)
            # 집계량1[new_index].to_excel(writer, sheet_name = '일자별__검사현황', startrow=5)



            # 집계량2.reindex(columns=pivot_label_shift).to_excel(writer, sheet_name = '일자별_호기교대별_검사현황', startrow=5)
            # 집계량1[new_index].to_excel(writer, sheet_name = '일자별__검사현황', startrow=5)


        if _action == 'DB_CHK':
            try:
                sheet_Name = 'DB_CHK'
                self.WB2.sheets[sheet_Name].activate()

                집계량1 = pd.pivot_table(검사_df[cond_Total], index='일자', columns = ['호기','교대','반'], values=['일보id'],
                                        margins=True, margins_name='합계', aggfunc="count")

                iterables = [['일보id'], ['2호기', '3호기'], ['조간', '석간', '야간'],[5,6,7,8,9,10]]
                pivot_label_shift = pd.MultiIndex.from_product(iterables, names=['일보id', '호기','교대','반'])

                집계량1_reIndex = 집계량1.reindex(columns=pivot_label_shift)

                self.WB2.sheets[sheet_Name].range('a1').value = self.Base_Date.date().toPyDate()
                self.WB2.sheets[sheet_Name].range('a4:bd10000').clear_contents()
                self.WB2.sheets[sheet_Name].range('a4').value = 집계량1_reIndex
                # sht = self.WB2.sheets['test'].activate
                # sht.app.Copy(Before=sht.app)

                # 집계량1_reIndex.to_excel(df_to_XL_PRINT), sheet_name ='DB_CHK2', startrow=4)
            except Exception as e:
                alert(str(e), '오류')


        self.WB2.save()


    def Combo_Length1_changed(self):
        if float(self.qt_Length1.currentText()) >= 2400:
            self.qt_Length2.setCurrentText(self.qt_Length1.currentText())
            self.qt_Length3.setCurrentText(self.qt_Length1.currentText())
            self.qt_Length4.setCurrentText(None)
        else:
            self.qt_Length2.setCurrentText(self.qt_Length1.currentText())
            self.qt_Length3.setCurrentText(self.qt_Length1.currentText())
            self.qt_Length4.setCurrentText(self.qt_Length1.currentText())


    def Delete_db(self):   #일보삭제
        try:
            # 기본정보 추출
            일자_date = self.Product_Date.date().toPyDate()
            일자 = 일자_date.strftime('%Y%m%d')
            호기 = self.qt_HOGI.currentText()
            날씨 = self.qt_Weather.currentText()
            교대 = self.qt_Shift.currentText()
            반 = int(self.qt_BAN.currentText())
            검사원 = self.qt_Shift.currentText()

            SaveFileName = 호기 + '_' + 일자 + '_'  + str(반) + '_' + 교대

            Q_delete = confirm('일보를 삭제할까요?', '확인', buttons=['예', '아니요'])

            if Q_delete=='예':
                # DB Setting ------
                conn = sqlite3.connect(DBNAME)

                # 생산data_df.set_index('일보id', inplace=True)
                del_sql_txt = "DELETE FROM 생산DataT WHERE 일보id = ?"

                cur = conn.cursor()
                cur.execute(del_sql_txt, (SaveFileName,))
                conn.commit()
                conn.close()

                os.remove(PICKLE_DIR+SaveFileName+'.pickle')
                os.remove(PICKLE_DIR+SaveFileName+'_sub.pickle')

                alert(title='삭제', text='삭제 완료')

        except Exception as e:
            alert(title='오류', text=str(e))


    def ReWrite_from_Pickle(self):   #검사일보 수정
        try:
            xl.Book(XL_DIR + '제품검사일보_Write.xlsx').app.quit()
            self.WB0 = xl.Book(XL_DIR + '제품검사일보_Base.xlsx')
            self.WB0.save(XL_DIR + '제품검사일보_Write.xlsx')

            일자 = self.Product_Date.date().toPyDate()
            호기 = self.qt_HOGI.currentText()
            날씨 = self.qt_Weather.currentText()
            교대 =self.qt_Shift.currentText()
            반 = self.qt_BAN.currentText()
            검사원 = self.qt_Worker.currentText()

            SaveFileName = 호기 + '_' + 일자.strftime('%Y%m%d') + '_'  + 반 + '_' + 교대

            df = pd.read_pickle(PICKLE_DIR + SaveFileName + '.pickle')
            self.WB0.sheets['일보'].range('b9').value = df.loc[:, :'샘플'].values
            self.WB0.sheets['일보'].range('b4').value = df.loc[0,'일자']
            self.WB0.sheets['일보'].range('f4').value = df.loc[0,'호기']
            self.WB0.sheets['일보'].range('j4').value = df.loc[0,'날씨']
            self.WB0.sheets['일보'].range('m4').value = df.loc[0,'교대']
            self.WB0.sheets['일보'].range('o4').value = df.loc[0,'반']
            self.WB0.sheets['일보'].range('s4').value = df.loc[0,'검사원']

            self.WB0.sheets['일보'].range('ao6:au15').clear_contents()

            df2 = pd.read_pickle(PICKLE_DIR + SaveFileName + '_sub.pickle')
            self.WB0.sheets['일보'].range('an6').value = df2.values

            일보_Data = np.array(self.WB0.sheets['일보'].range(R0).value)
            iRow_CHK = 일보_Data[:,2:4]
            일보_Row = 9

            for icount, chk in enumerate(일보_Data[:,3:7]):   #규격행 체크

                소로트체크_List = [k for k in chk if k != None]

                if np.sum(소로트체크_List) > 1500:
                    self.WB0.sheets['일보'].range('b' + str(icount+일보_Row) + ":am" + str(icount + 일보_Row)).api.Interior.ColorIndex = 15

            self.WB0.save()

            alert(title='일보수정', text='일보읽어오기 완료')

        except Exception as e:
            self.WB0.app.quit()
            alert(title='오류', text=str(e))


    def Job_Change(self):   #규격변경

        일보_Data = np.array(self.WB0.sheets['일보'].range(R0).value)
        일보_Row = 9
        chk_OK = None
        icount = 1

        selection_cell = self.WB0.app.selection

        입력행선택 = confirm('규격을 입력할 행을 선택하셨나요?', '확인', buttons=['예', '아니요'])

        if (9<= selection_cell.row <=55) and (입력행선택=='예') and (np.sum(일보_Data[selection_cell.row-일보_Row, 3]) == None):
            품종 = self.qt_PJ.currentText()
            두께 = self.qt_Thick.currentText()
            나비 = self.qt_Width.currentText()
            길이1 = self.qt_Length1.currentText()
            길이2 = self.qt_Length2.currentText()
            길이3 = self.qt_Length3.currentText()
            길이4 = self.qt_Length4.currentText()
            self.WB0.sheets['일보'].range('b' + str(selection_cell.row)).value = np.array([품종, 두께, 나비, 길이1, 길이2, 길이3, 길이4])

        else: alert(title='오류', text = '규격을 입력할 행 선택이 잘 되었는지 확인하세요')

        self.WB0.sheets['일보'].activate()


    def Load_xl_Form(self):  # 검사일보 추가
        xl.Book(XL_DIR + '제품검사일보_Write.xlsx').close()
        self.WB0 = xl.Book(XL_DIR + '제품검사일보_Base.xlsx')
        self.WB0.save(XL_DIR + '제품검사일보_Write.xlsx')


    def Save_to_Pickle(self): #검사일보 저장
        try:
            일보_Columns = ['No',	'시각',	'규격_나비', '길이_가',	'길이_나',	'길이_다',	'길이_라', '판정_가',	'판정_나',	'판정_다',	'판정_라',	'대각',	'측각_좌',	'측각_우',
                        '길이1', '길이2',	'길이3',	'길이4', '나비',	'두께',	'내박리성_표면',	'내박리성_이면',	'함수율',	'BF_MD',	'BF_CD',
                        '무게',	'비중',	'MPSM',	'전흡수율',	'타점',	'요철',	'슬러리',	'각불량',	'외관',	'변파',	'접착',	'기포',	'샘플']  # 38개


            일보_Data =     np.array(self.WB0.sheets['일보'].range(R0).value)

            일보_Data[:,30:][일보_Data[:,30:]==None] = 0.

            규격Row_CHK=[]
            생산규격 = []
            Range_Lot = []
            생산Data = []
            위치_List = ['가','나','다','라']


            소로트 = 1
            일보_Row = 9

            iRow_CHK = 일보_Data[:,2:4]

            # 기본정보 추출
            일자_date = self.WB0.sheets['일보'].range('b4').value
            일자 = 일자_date.strftime('%Y%m%d')
            호기 = self.WB0.sheets['일보'].range('f4').value
            날씨 = self.WB0.sheets['일보'].range('j4').value
            교대 = self.WB0.sheets['일보'].range('m4').value
            반 = int(self.WB0.sheets['일보'].range('o4').value)
            검사원 = self.WB0.sheets['일보'].range('s4').value

            print(일자, 호기, 날씨, 교대, 반, 검사원)
            SaveFileName = 호기 + '_' + 일자 + '_'  + str(반) + '_' + 교대


            for icount, chk in enumerate(일보_Data[:,3:7]):   #규격행 체크

                소로트체크_List = [k for k in chk if k != None]

                if np.sum(소로트체크_List) > 1500:
                    규격Row_CHK.append(icount)

                    규격_Row = icount

                    품종 = 일보_Data[규격_Row, 0]
                    두께 = float(일보_Data[규격_Row, 1])
                    나비 = float(일보_Data[규격_Row, 2])
                    self.WB0.sheets['일보'].range('b' + str(icount+일보_Row) + ":am" + str(icount + 일보_Row)).api.Interior.ColorIndex = 15

                    소로트 = 1
                else:
                    if np.sum(소로트체크_List)>0:
                        self.WB0.sheets['일보'].range('b' + str(icount+일보_Row)).value = 소로트

                        for k, GSR in enumerate(소로트체크_List):
                            위치 = 위치_List[k]
                            길이 = 일보_Data[규격_Row, k+3]
                            검사량 = GSR
                            판정 = 일보_Data[icount,k+7]
                            치수_길이 = 일보_Data[icount,k+14]
                            if 판정=='M': 미선별 = 검사량
                            else:미선별 = 0

                            if k==0:
                                # if 일보_Data[icount, 1] != None:
                                #     시각 = 일보_Data[icount, 1]
                                #     if len((시각))<4:
                                #         시각 = 일보_Data[icount-1, 1][0:2] + 시각
                                #         self.WB0.sheets['일보'].range('c' + str(icount+일보_Row)).value = 시각
                                시각 =         일보_Data[icount, 1]
                                창고위치 =     일보_Data[icount, 2]
                                대각 =         일보_Data[icount,11]
                                측각_좌 =      일보_Data[icount,12]
                                측각_우 =      일보_Data[icount,13]
                                치수_나비 =    일보_Data[icount,18]
                                치수_두께 =    일보_Data[icount,19]
                                내박리성_표면, 내박리성_뒷면, 함수율, BF_MD, BF_CD, 무게 = 일보_Data[icount,20:26]

                                if (무게 != None) and (두께*나비*길이) > 0:
                                    비중 = 무게/(두께*나비*길이)*1000000
                                    MPSM = 무게/(나비*길이)*1000000
                                    self.WB0.sheets['일보'].range('ab' + str(icount+일보_Row)).value = 비중
                                    self.WB0.sheets['일보'].range('ac' + str(icount+일보_Row)).value = MPSM
                                else:
                                    비중 = None
                                    MPSM = None

                                전흡수율 = 일보_Data[icount,28]
                                타점 =     일보_Data[icount,29]

                                요철, 슬러리, 각불량, 외관, 변파, 접착, 기포, 샘플 = 일보_Data[icount,30:38]

                            else:
                                대각, 측각_좌, 측각_우, 치수_나비, 치수_두께, 내박리성_표면, 내박리성_뒷면, 함수율, BF_MD, BF_CD, \
                                무게, 비중, MPSM, 전흡수율, 타점, 요철, 슬러리, 각불량, 외관, 변파, 접착, 기포, 샘플 = [None]* 23

                            생산Data.append([일자, 호기, 날씨, 교대, 반, 검사원, 품종, 두께, 나비, 길이, 소로트, 시각, 위치, 창고위치, 검사량, 미선별,
                                            판정, 치수_길이, 대각, 측각_좌, 측각_우, 치수_나비, 치수_두께, 내박리성_표면, 내박리성_뒷면,
                                            함수율, BF_MD, BF_CD, 무게, 비중, MPSM, 전흡수율, 타점, 요철, 슬러리, 각불량, 외관,
                                            변파, 접착, 기포, 샘플])

                        소로트 += 1


            생산data_Columns = ['일자', '호기', '날씨', '교대', '반', '검사원', '품종', '두께', '나비', '길이', '소로트','시각', '위치','창고위치', '검사량', '미선별',
                                '판정','치수_길이', '대각', '측각_좌', '측각_우', '치수_나비', '치수_두께', '내박리성_표면', '내박리성_뒷면',
                                '함수율', 'BF_MD', 'BF_CD', '무게', '비중', 'MPSM', '전흡수율', '타점', '요철', '슬러리', '각불량', '외관',
                                '변파', '접착', '기포', '샘플']

            생산data_df = pd.DataFrame(np.array(생산Data), columns=생산data_Columns)
            생산data_df['불량계'] = 생산data_df.요철 + 생산data_df.슬러리 + 생산data_df.각불량 + 생산data_df.외관 + 생산data_df.변파 + 생산data_df.접착 + 생산data_df.기포
            생산data_df['환산계수'] = (생산data_df.두께/9.5 * 생산data_df.나비/900 * 생산data_df.길이/1800) * 0.5
            생산data_df['환산계수'] = 생산data_df['환산계수'].apply(lambda x: round(x, 3))
            생산data_df['일보id'] = SaveFileName
            # 생산data_df = 생산data_df.fillna(0)

            insert_sql_txt = "INSERT INTO 생산DataT(일자, 호기, 날씨, 교대, 반, 검사원, 품종, 두께, 나비, 길이, 소로트, 시각, 위치, 검사량, 미선별, \
                                            판정, 치수_길이, 대각, 측각_좌, 측각_우, 치수_나비, 치수_두께, 내박리성_표면, 내박리성_뒷면,\
                                            함수율, BF_MD, BF_CD, 무게, 비중, MPSM, 전흡수율, 타점, 요철, 슬러리, 각불량, 외관,\
                                            변파, 접착, 기포, 샘플, 불량계, 환산계수, 일보id) \
                              VALUES(?,?,?,?,?,?,?,?,?,?, ?,?,?,?,?,?,?,?,?,?, ?,?,?,?,?,?,?,?,?,?, ?,?,?,?,?,?,?,?,?,?, ?,?)"  # 42개

            # 규격별 집계  ---
            집계규격 = []
            집계_품종 = []
            집계_두께 = []
            집계_나비 = []
            집계_길이 = []

            for m in 규격Row_CHK:
                집계_품종 = 일보_Data[m, 0]
                집계_두께 = float(일보_Data[m, 1])
                집계_나비 = int(일보_Data[m, 2])
                집계_길이 = [int(k) for k in 일보_Data[m, 3:7] if k != None]
                집계규격 += [[집계_품종, 집계_두께, 집계_나비, k] for k in 집계_길이]

            집계규격_List = sorted(list(set(map(tuple,집계규격))))

            규격별합계_List = []
            일보집계_col = 41
            생산평 = 정품평 = 불량평 = 샘플평 = 미선별 = 0.


            for PJ, thick, width, length in 집계규격_List:
                cond = (생산data_df.품종==PJ) & (생산data_df.두께==thick) & (생산data_df.나비==width) & (생산data_df.길이==length)
                규격별합계_df = 생산data_df[cond].loc[:, ['검사량','미선별','불량계','샘플']].sum()
                검사H, 미선H, 불량H, 샘플H = 규격별합계_df.values
                정품H = 검사H - 미선H - 불량H - 샘플H
                self.WB0.sheets['일보'].range(6,일보집계_col).value = np.array([PJ, thick, width,length, 검사H, 정품H, 불량H, 샘플H, 미선H]).reshape(9,1)
                # self.WB0.sheets['일보'].range(10,일보집계_col).value = np.array([length, 검사H, 정품H, 불량H, 샘플H, 미선H]).reshape(6,1)

                환산H = np.round((thick/9.5 * width/900 * length/1800 * 0.5), 3)

                생산평 += 검사H * 환산H
                정품평 += 정품H * 환산H
                불량평 += 불량H * 환산H
                샘플평 += 샘플H * 환산H
                미선별 += 미선H * 환산H

                일보집계_col += 1

            #평수 집계-----
            self.WB0.sheets['일보'].range('ao41').value = np.array([생산평, 정품평, 불량평, 샘플평, 미선별]).reshape(5,1)

            if os.path.exists(PICKLE_DIR+SaveFileName+'.pickle'):
                save_chk = confirm('동일한 파일이 존재 합니다. 저장할까요?', buttons=['예', '아니요'] )
            else: save_chk='예'

            if save_chk=='예':

                # Pickle data 저장 ----------------------------

                일보저장_Data = np.array(self.WB0.sheets['일보'].range(R0).value)
                일보_Data_sub = np.array(self.WB0.sheets['일보'].range('an6:au46').value)

                일보_df = pd.DataFrame(일보저장_Data, columns=일보_Columns)
                일보_df['일자'] = 일자_date
                일보_df['호기'] = 호기
                일보_df['날씨'] = 날씨
                일보_df['교대'] = 교대
                일보_df['반'] = 반
                일보_df['검사원'] = 검사원

                일보_sub_df = pd.DataFrame(일보_Data_sub)

                일보_df.to_pickle(PICKLE_DIR + SaveFileName + '.pickle')
                일보_sub_df.to_pickle(PICKLE_DIR + SaveFileName + '_sub.pickle')

                # DB 저장 ---------------------------------------
                # DB Setting ------
                conn = sqlite3.connect(DBNAME)

                # 생산data_df.set_index('일보id', inplace=True)
                del_sql_txt = "DELETE FROM 생산DataT WHERE 일보id = ?"

                cur = conn.cursor()
                # 생산data_df.to_sql('생산DataT', con=conn, if_exists='replace', index=False)
                cur.execute(del_sql_txt, (SaveFileName,))
                # cur.executemany(insert_sql_txt, 생산data_df.values)
                생산data_df.to_sql('생산DataT', con=conn, if_exists='append', index=False)
                conn.commit()
                conn.close()

                alert(title='데이터 저장', text='정상적으로 데이터가 저장되었습니다.')

            self.WB0.save()
            self.WB0.app.quit()

        except Exception as e:
            alert(title='오류', text=str(e) + '초기값(일자 등) 확인')


    def Read_from_Pickle(self):  #직전 검사일보 불러오기
        try:
            # xl.Book(XL_DIR + '제품검사일보_temp.xlsx').app.quit()
            self.WB0 = xl.Book(XL_DIR + '제품검사일보_Write.xlsx')
            # self.WB0.save(XL_DIR + '제품검사일보_Read.xlsx')
            alert(title='완료', text='직전일보 불러오기 완료')

        except Exception as e:
            alert(title='오류', text=str(e))


    def Input_xl_Init_Value(self):  #일보 초기값 입력
        try:
            # 초기값 입력 -------
            self.WB0.sheets['일보'].range('b4').value = self.Product_Date.date().toPyDate()
            self.WB0.sheets['일보'].range('f4').value = self.qt_HOGI.currentText()
            self.WB0.sheets['일보'].range('j4').value = self.qt_Weather.currentText()
            self.WB0.sheets['일보'].range('m4').value =self.qt_Shift.currentText()
            self.WB0.sheets['일보'].range('o4').value = self.qt_BAN.currentText()
            self.WB0.sheets['일보'].range('s4').value = self.qt_Worker.currentText()

            품종 = self.qt_PJ.currentText()
            두께 = self.qt_Thick.currentText()
            나비 = self.qt_Width.currentText()
            길이1 = self.qt_Length1.currentText()
            길이2 = self.qt_Length2.currentText()
            길이3 = self.qt_Length3.currentText()
            길이4 = self.qt_Length4.currentText()

            self.WB0.sheets['일보'].range('b9').value = np.array([품종,두께, 나비, 길이1, 길이2, 길이3, 길이4])

            self.WB0.app.activate(steal_focus=True)
            self.WB0.save()
            alert(title='완료', text='초기값 입력 완료')
        except Exception as e:
            alert(title='오류', text=str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(True)
    main_win = GI_Main()
    main_win.show()
    sys.exit(app.exec_())
    pass