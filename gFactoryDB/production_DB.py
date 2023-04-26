import xlwings as xl
import numpy as np
import pandas as pd
import pymysql
import mysql.connector as mdb

# import win32com.client
from xlwings.utils import rgb_to_int
from pymsgbox import alert, confirm, password, prompt
from xlwings.utils import rgb_to_int
from pathlib import Path
from pymsgbox import alert, confirm, password, prompt
from sqlalchemy import create_engine
from xlwings import Chart, Shape

# DIR_ROOT = Path(__file__)
XLS = 'd:\\G_FactoryDB_Asset\\XLS\\'
WB_prodDB = xl.Book(XLS + 'ProductionDB.xlsb')
SH_ProdT = WB_prodDB.sheets['Prod_T']

# DB Source ------------------
USER = 'leekr2'
HOST = '10.50.3.163'
PASSWORD = 'g1234'
DB = 'gfactoryDB'


def ConnectDB():
    
    # DB Setting ------
    try:
        # Get Cursor -------------------------------
        # pymysql.install_as_MySQLdb()
        
        engine = create_engine("mysql://{user}:{password}@{host}/{db}".format(user=USER, password=PASSWORD, host=HOST, db=DB))
        # conn = engine.connect(host=HOST, user=USER, password=PASSWORD, db=DB, charset='utf8')
        conn = pymysql.connect(host=HOST, user=USER, password=PASSWORD, database=DB, charset='utf8')
        cur = conn.cursor()        
        conn.commit()        
        
        
    except mdb.Error as e:
        print(f"Error connecting to MySQL Platform: {e}")
        sys.exit(1)
        
    return engine, conn, cur



def DataUpload():  
    
    [engine, conn, cur] = ConnectDB()
    
    count_xl_data = SH_ProdT.range('c8').value
    rowN = int(11 + count_xl_data - 1)

    
    df_prod_T = pd.DataFrame(SH_ProdT.range('a11:ay' + str(rowN)).value, columns=SH_ProdT.range('a9:ay9').value)
    df_prod_T['KeyInCode'] = SH_ProdT.range('b8').value
  
    print(df_prod_T)
    
    curCHK = cur.execute("select * from Prod_T where KeyInCode=%s", str(SH_ProdT.range('b8').value))
    conn.commit()

    # DB upload-------------------------------------------------------
    try:
        # Delete -----
        if curCHK>0: 
            CHK_Replace = confirm('Data Exist. Surely Replace them?', 'Confirm Replace Data', buttons=['Yes, Sure', 'No'])
            
            if CHK_Replace=='Yes, Sure':    
                sqltxt_del_1 = "DELETE FROM Prod_T WHERE KeyInCode=%s"
                cur.execute(sqltxt_del_1, str(SH_ProdT.range('b8').value))
                conn.commit()        
                # Update Data ------------------
                df_prod_T.to_sql('Prod_T', con=engine, if_exists='append', index=False)
                # conn.commit()
                conn.close()
                print('Applied', CHK_Replace)
            else:
                print('canceled')
            
    except Exception as e:
        alert(e)

if __name__ == '__main__':
    import sys
    # selectPro = int(sys.argv[1])    
    DataUpload()

    # elif selectPro == 2: ConnectDB(USER=USER, HOST=HOST, PASSWORD=PASSWORD, DB=DB)
    sys.exit()
