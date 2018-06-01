# coding=utf-8
import os, re, sys, time, xlrd, pyodbc, datetime
import win32com.client as win32
import numpy as np
import pandas as pd
import myfun

### function generate email, then send or show based on parameter 'auto'
def CreateEmail(recipient, subject, msg, isattach, attachloc):
	#print 'Generate'
	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	mail.SentOnBehalfOfName = 'Business Reporting (IG)' #'shcomig@investorsgroup.com'
	mail.To = recipient
	#mail.Bcc = 'Gail.Purcell@investorsgroup.com;Corinne.Fontaine@investorsgroup.com'
	mail.Subject = subject
	mail.HTMLBody = msg
	if isattach == 1:
		mail.Attachments.Add(Source = attachloc)
		
	mail.Close(0)
		
### template for Original allocation to PIP or SOP only		
template1 = '''Sent on behalf of Corinne Fontaine, Manager, Consultant Compensation and Reporting<br/><br/>

Your April 15, 2018 commission statement contains two ASF Advance payment recovery deductions because you were not deducted on your February 15, 2018 commission statement.<br/><br/>
Please accept my apologies for any inconvenience this may cause.
'''		

template2 = u'''Message envoyé au nom de Corinne Fontaine, directrice, Rémunération des conseillers et information d’affaires<br/><br/> 

Votre relevé de commissions du 15 avril 2018 comprend deux déductions pour le recouvrement d'avance de CSA parce que la déduction n’a pas été faite comme prévu sur votre relevé de commissions du 15 février 2018.<br/><br/>
Veuillez accepter mes excuses pour les inconvénients que cette erreur a pu vous causer.
'''		

rlist1 = []
sublist1 = []
msglist1 = []
isattach1 = []
filelist1 = []

driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
db_file = r"F:\\Files For\\West Wang\\Rates.accdb;"
user = "admin"
password = ""
odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)

sql = '''
SELECT DISTINCT 
	[Tem].[Cslt]
	,BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_STATUS
	,BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_TERM_DTE
	,BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_LANGUAGE
	,BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_EMAIL_ALIAS
FROM [Tem]
LEFT JOIN BRANUSER_BRAN_LKG_CSLT_CURR ON [Tem].[Cslt] = BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_NUM
WHERE BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_STATUS = 'Active';
'''

conn = pyodbc.connect(odbc_conn_str)
cur = conn.cursor()
df = pd.read_sql_query(sql,conn)

writer = pd.ExcelWriter('ASFEmailList.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
#sys.exit('stop here')

for index, row in df.iterrows():
	if row['LKG_CSLT_LANGUAGE'] == 'E':
		rlist1.append(row['LKG_CSLT_EMAIL_ALIAS'])
		#rlist1.append('west.wang@investorsgroup.com')
		sublist1.append(u'''Recovery of ASF Advances''')
		#sublist1.append('This is E test' + row['Email'])
		msglist1.append(template1)
		isattach1.append(0)
		filelist1.append('')		
	else:
		rlist1.append(row['LKG_CSLT_EMAIL_ALIAS'])
		#rlist1.append('west.wang@investorsgroup.com')
		sublist1.append(u'''Récupération des avances de CSA''')
		#sublist1.append('This is E test' + row['Email'])
		msglist1.append(template2)
		isattach1.append(0)
		filelist1.append('')			
		

for e, s, m, a, f in zip(rlist1, sublist1, msglist1, isattach1, filelist1):
	CreateEmail(e, s, m, a, f)