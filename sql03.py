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
#db_file = r"F:\3-Compensation Programs\Net Flows Bonus\NetFlow.accdb;"
db_file = r"F:\Files For\Hai Yen Nguyen\RO Reserve Fund\RO Reserve Fund.accdb;"
user = "wwang"
password = "west33"
odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)

#sql = '''
#SELECT LKG_CSLT_SMPL_DTE, LKG_CSLT_NUM, LKG_CSLT_STATUS, LKG_CSLT_POSITION, LKG_CSLT_EMAIL_ALIAS FROM BRANUSER_BRAN_LKG_CSLT_CURR
#'''
#sql = sql % (values)

#sql = '''
#		SELECT C.LKG_CSLT_RO_NUM AS RO
#			,C.LKG_CSLT_NUM, C.LKG_CSLT_NAM_FULL, C.LKG_CSLT_SMPL_DTE 
#		FROM BRANUSER_BRAN_LKG_CSLT AS C
#			WHERE C.LKG_CSLT_SMPL_DTE > # 10/01/2017 # AND C.LKG_CSLT_SMPL_DTE < # 01/01/2018 #
#			AND C.LKG_CSLT_POSITION = 'REGIONAL DIRECTOR'
#			AND C.LKG_CSLT_STATUS = 'Active' AND C.LKG_CSLT_NAM_FULL <> 'INTERIM RD, INTERIM'
#'''
sql = '''
SELECT
tCmsn_Bal.cmsn_cslt_num AS CSLT, 
tCmsn_Bal.cmsn_acct_type,
tCmsn_Bal.cmsn_tot_bal
FROM tCmsn_Bal
WHERE 
cmsn_bal_dte = # 12/31/2017 # AND tCmsn_Bal.cmsn_cslt_num < 80000

'''

conn = pyodbc.connect(odbc_conn_str)
cur = conn.cursor()
df1 = pd.read_sql_query(sql,conn)

db_file = r"F:\Files For\West Wang\Rates.accdb;"
user = "wwang"
password = "west33"
odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)

sql = '''
SELECT DISTINCT
	BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_NUM AS CSLT
	,BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_NAM_FULL
	,BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_RO_NUM
	,BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_POSITION
	,BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_STATUS
FROM BRANUSER_BRAN_LKG_CSLT
WHERE (
		((BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_SMPL_DTE) = #12/31/2017#)
		);
'''
conn = pyodbc.connect(odbc_conn_str)
cur = conn.cursor()
df2 = pd.read_sql_query(sql,conn)

df2 =  df2.merge(df1, left_on='CSLT', right_on='CSLT', how='inner')
#df.to_csv('sql.csv', index=False)
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('name.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df2.to_excel(writer, index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()