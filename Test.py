print "Hello World"

import os
import os
import re
import fnmatch
import pandas as pd
import pyodbc

#df = pd.read_excel('F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMA Daily 1.xlsx', sheet_name='SMA')
#print df.head()
#for index, row in df.iterrows():
	#print tuple(row)
driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
db_file = r"F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMA.accdb;"
user = "admin"
password = ""
odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)

df = pd.read_excel('F:\\3-Compensation Programs\\IIROC Compensation\\CMMPN.XS.ODS.DAILY.SMA.EVENTS.20171117.xls', header=None)
df1 = (df.loc[df[0] == 'D'])
	
#if not df1.empty:
#	print df1
	#SMAdata = SMAdata.append(df1, ignore_index=True)

table = "SMA"
columns = "'Event Record Type', 'Event Effective Date', 'Event Process Date', 'Event Activity Type', 'Event Activity Description', 'Event Gross Amount', 'Plan Product Code', 'Account Market Value', 'Client Number', 'Client Last Name', 'Client Given Name', 'Client Servicing Consultant Number', 'Client Deceased Indicator', 'Client Company Name', 'Client Province Code', 'Account Number', 'Account Dealer Code', 'Account IGSI Net Share Quantity', 'Product Code', 'Product Share Price Amount', 'Product IGSI Symbol', 'Product Description', 'Product Security Type', 'Product Security Class', 'Product Security Category'"

for row in df1.to_records(index=False):
	values = ", ".join(['\'%s\'' % x for x in row])
	values = values.replace("'nan'", "NULL")
	#print values
	sql = '''INSERT INTO %s VALUES (%s, '20177777');'''
	sql = sql % (table, values)
	print sql
	conn = pyodbc.connect(odbc_conn_str)
	cursor = conn.cursor()
	cursor.execute(sql)
	cursor.commit()
	cursor.close()
conn.close()  	