print "Hello World"

# -*- coding: utf-8 -*-
import pyodbc
driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
db_file = r"C:\pycode\PillarAward.accdb;"
#db_file = r"F:\Files For\Hai Yen Nguyen\IIROC reporting\IIROC_Income.accdb;"
user = "admin"
password = ""
odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)

#sql = "SELECT TOP 10 LKG_CSLT_NUM, LKG_CSLT_NAM_FULL FROM BRANUSER_BRAN_LKG_CSLT"
#sql = "SELECT * FROM Cslt"
sql = "INSERT INTO Cslt VALUES (6, '004', 'Tom', 1)"

conn = pyodbc.connect(odbc_conn_str)
cur = conn.cursor()
#for row in cur.tables():
    #print row.table_name
#	print row, type(row)

cur.execute(sql)
#rows = cur.fetchall()
#for row in rows:
#    print row
	
cur.close()
conn.close()

print 'finished'