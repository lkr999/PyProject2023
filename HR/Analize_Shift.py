
import xlwings as xl
import numpy as np
import pandas as pd
import pymysql

from pymsgbox import alert, confirm, password, prompt
from xlwings.utils import rgb_to_int
from pathlib import Path
from pymsgbox import alert, confirm, password, prompt
from sqlalchemy import create_engine

# DIR_ROOT = Path(__file__)
XLS = 'd:\\G_FactoryDB_Asset\\XLS_Local\\'
WB0 = xl.Book(XLS + 'HR_analyze.xlsb')
SH_ProdT = WB0.sheets['Prod_T']

