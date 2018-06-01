#print "Hello World"
#
## -*- coding: utf-8 -*-
#import pypyodbc
#driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
##db_file = r"C:\pycode\PillarAward.accdb;"
#db_file = r"F:\Files For\Hai Yen Nguyen\IIROC reporting\IIROC_Income.accdb;"
#user = "admin"
#password = ""
#odbc_conn_str = r"Driver={};Dbq={};".format(driver, db_file)
#
#sql = "SELECT TOP 10 LKG_CSLT_NUM, LKG_CSLT_NAM_FULL FROM BRANUSER_BRAN_LKG_CSLT"
#pypyodbc.lowercase = False
#conn = pypyodbc.connect(odbc_conn_str)
#cur = conn.cursor()
#for row in cur.tables():
#    print row
##cur.execute("INSERT INTO Cslt VALUES (7, '004', 'Tom', 1)");
##cur.execute("SELECT * FROM Cslt");
#cur.execute(sql);
##while True:
##    row = cur.fetchone()
##    if row is None:
##        break
##    print(u"Consltant with name {0} is {1}".format(row.get("CsltName"), row.get("CsltType")))
#cur.close()
#conn.close()
#print 'finished'
import pandas as pd
import pyodbc
import myfun
import datetime 

cycdate = datetime.datetime.strptime(str('2017-10-31'), '%Y-%m-%d')

l = [15932, 5285, 14481, 12069, 11896, 15932, 12931, 18605, 12911] 
d = pd.DataFrame(l, columns=['id'])
enddate = datetime.datetime.strptime(str('2017/11/15'), '%Y/%m/%d')
sql = '''
		SELECT * FROM [qry_12_SalesBonus_Transitional vs Ongoing]
		'''
driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
db_file = r"F:\3-Compensation Programs\Our Clients, Our Future 2016\Sales Bonus\Accrual Sales Bonus estimate\Sales Bonus Premium Accrual.accdb;"
user = "admin"
password = ""
odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)
#sql = sql.format(tuple(d['id']))
conn = pyodbc.connect(odbc_conn_str)
cur = conn.cursor()
df = pd.read_sql_query(sql,conn)
df.to_csv('sales.csv', index=False)
