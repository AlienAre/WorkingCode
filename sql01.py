import os
import re
import sys
import time
from datetime import date
import fnmatch
import pandas as pd
import itertools as it
from openpyxl import load_workbook
import xlrd
import myfun as dd
import pyodbc
import datetime

#df = pd.read_excel('F:\\Files For\\West Wang\\Ad Hoc\\adsr-2241-results.xlsx')
#print(df.head())
#sys.exit("done")
#values = ", ".join(['\'%s\'' % x for x in df['Cslt']])
#values = ", ".join(['%s' % x for x in df['Cslt']])
driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
db_file = r"F:\Files For\West Wang\Rates.accdb;"
user = "admin"
password = ""
odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)

#sql = '''
#SELECT LKG_CSLT_SMPL_DTE, LKG_CSLT_NUM, LKG_CSLT_STATUS, LKG_CSLT_POSITION, LKG_CSLT_EMAIL_ALIAS FROM BRANUSER_BRAN_LKG_CSLT_CURR
#'''
#sql = sql % (values)

#sql = '''
#SELECT BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_NUM
#	,BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_NAM_LAST
#	,BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_NAM_FIRST
#	,BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_NAM_FULL
#	,BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_STATUS
#	,BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_EMAIL_ALIAS
#	,T5008.*
#FROM BRANUSER_BRAN_LKG_CSLT_CURR
#RIGHT JOIN T5008 ON T5008.Last = BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_NAM_LAST AND T5008.First = BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_NAM_FIRST
#'''

#sql = '''
#SELECT BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_SMPL_DTE
#	,BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_NUM
#	,BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_RO_NUM
#	,BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_STATUS
#	,SUM(IIF(tCslt_Debt_Bal.debt_type = "CMSN", tCslt_Debt_Bal.debt_amt, 0)) AS [CMSNBal]
#	,SUM(IIF(tCslt_Debt_Bal.debt_type = "AP", tCslt_Debt_Bal.debt_amt, 0)) AS [APBal]
#FROM BRANUSER_BRAN_LKG_CSLT
#INNER JOIN tCslt_Debt_Bal ON (BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_NUM = tCslt_Debt_Bal.debt_cslt_num)
#	AND (BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_SMPL_DTE = tCslt_Debt_Bal.debt_dte)
#WHERE BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_SMPL_DTE = # 12/31/2017 #
#GROUP BY
#BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_SMPL_DTE
#	,BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_NUM
#	,BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_RO_NUM
#	,BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_STATUS
#'''
sql = '''
SELECT * FROM qry_RODebt
'''

#------------------------
#f = open('test.txt', 'w')
#f.write(sql)
#f.close()
#sys.exit("done")
#-------------------------

conn = pyodbc.connect(odbc_conn_str)
cur = conn.cursor()
df1 = pd.read_sql_query(sql,conn)

#df =  df.merge(df1, left_on='Cslt', right_on='LKG_CSLT_NUM', how='left')
#df.to_csv('sql.csv', index=False)
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('name.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df1.to_excel(writer, index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()