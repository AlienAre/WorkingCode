import os, re, sys, time, xlrd, pyodbc, fnmatch
from datetime import date
import pandas as pd
import itertools as it
from openpyxl import load_workbook
import myfun as dd
import datetime

print 'Process date is ' + str(time.strftime("%m/%d/%Y"))
print 'Please enter the cycle end date (mm/dd/yyyy) you want to process:'

#-----------------------------------------------------
#------- get cycle date ----------------------
getcycledate = datetime.datetime.strptime(raw_input(), '%m/%d/%Y')
endday = getcycledate
startday = dd.getCStartDate(getcycledate)

print 'Cycle start date is ' + str(startday)
print 'Cycle end date is ' + str(endday)

r24startdate = dd.get24CycleStartDate(endday)
r24enddate = dd.get24CycleEndDate(endday)

print 'Rolling 24 Cycle start date is ' + str(r24startdate)
print 'Rolling 24 Cycle end date is ' + str(r24enddate)
	
#sys.exit("done")
#values = ", ".join(['\'%s\'' % x for x in df['Cslt']])
#values = ", ".join(['%s' % x for x in df['Cslt']])
driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
#db_file = r"F:\3-Compensation Programs\Net Flows Bonus\NetFlow.accdb;"
db_file = r"F:\3-Compensation Programs\Achievement Level Compensation Differential\NewPro\Advance AL.accdb;"
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
BRANUSER_BRAN_SALES_CRED.SCRED_CSLT_NUM AS CSLT, 
SUM(BRANUSER_BRAN_SALES_CRED.SCRED_CYC_TTL_RECOG) AS AMT
FROM BRANUSER_BRAN_SALES_CRED
WHERE 
BRANUSER_BRAN_SALES_CRED.SCRED_SMPL_DTE >= # ''' + r24startdate.strftime('%m/%d/%Y') + ''' # AND BRANUSER_BRAN_SALES_CRED.SCRED_SMPL_DTE <= # ''' + r24enddate.strftime('%m/%d/%Y') + ''' #
GROUP BY
BRANUSER_BRAN_SALES_CRED.SCRED_CSLT_NUM 
HAVING 
SUM(BRANUSER_BRAN_SALES_CRED.SCRED_CYC_TTL_RECOG) > 0
'''
#print sql

conn = pyodbc.connect(odbc_conn_str)
cur = conn.cursor()
df1 = pd.read_sql_query(sql,conn)

sql = '''
SELECT DISTINCT
	qry_CsltListForAL.LKG_CSLT_NUM AS CSLT
	,qry_CsltListForAL.LKG_CSLT_NAM_FULL
	,qry_CsltListForAL.LKG_CSLT_RO_NUM
	,qry_CsltListForAL.LKG_CSLT_SALES_START_DTE
	,'' AS TermDate
	,qry_CsltListForAL.LKG_CSLT_POSITION
	,qry_CsltListForAL.ACHV_PAID_COMP_AL
	,qry_CsltListForAL.LKG_CSLT_SDLR_NUM
FROM qry_CsltListForAL
WHERE (
		((qry_CsltListForAL.LKG_CSLT_SMPL_DTE) = # ''' + endday.strftime('%m/%d/%Y') + ''' #)
		);
'''
conn = pyodbc.connect(odbc_conn_str)
cur = conn.cursor()
df2 = pd.read_sql_query(sql,conn)

df2 =  df2.merge(df1, left_on='CSLT', right_on='CSLT', how='inner')
#df.to_csv('sql.csv', index=False)
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('AdvAL.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df2.to_excel(writer, index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()