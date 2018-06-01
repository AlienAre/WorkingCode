print "Hello World"

import os, re, sys
import fnmatch
import pandas as pd
import pyodbc

driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
db_file = r"F:\\3-Compensation Programs\\IIROC Compensation\\SMA, FBA Compensation\\SMA.accdb;"
user = "admin"
password = ""
odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)

#df = pd.read_excel('F:\\3-Compensation Programs\\IIROC Compensation\\CMMPN.XS.ODS.DAILY.SMA.EVENTS.20171117.xls', header=None)
#df1 = (df.loc[df[0] == 'D'])
df = pd.read_excel('C:\\pycode\\Book2.xlsx', header=None)

#print df.head()
df[15] = df[15].str.replace("'", "")
df[27] = df[27].str.replace("'", "")
#print df[15]
#sys.exit("done")

table = "tbl_SMA"
columns = '''
[CycleDate],
[TransType],
[Event Record Type],
[Event Effective Date],
[Event Process Date],
[Event Activity Type],
[Event Activity Description],
[Event Gross Amount],
[Plan Product Code],
[Account Market Value],
[Client Number],
[Client Last Name],
[Client Given Name],
[Client Servicing Consultant Number],
[Client Deceased Indicator],
[Client Company Name],
[Client Province Code],
[Account Number],
[Account Dealer Code],
[Account IGSI Net Share Quantity],
[Product Code],
[Product Share Price Amount],
[Product IGSI Symbol],
[Product Description],
[Product Security Type],
[Product Security Class],
[Product Security Category],
[Notes]
'''

for row in df.to_records(index=False):
	values = ", ".join(['\'%s\'' % x for x in row])
	values = values.replace("'nan'", "NULL")
	#print values
	sql = '''INSERT INTO %s (%s) VALUES (%s);'''
	sql = sql % (table, columns, values)
#-------------------------------------	
#	print sql
#	with open("Output.txt", "w") as text_file:
#		text_file.write(sql)
#	sys.exit("done")	
#--------------------------------------	
	conn = pyodbc.connect(odbc_conn_str)
	cursor = conn.cursor()
	cursor.execute(sql)
	cursor.commit()
	cursor.close()
	#sys.exit("done")
conn.close()  	