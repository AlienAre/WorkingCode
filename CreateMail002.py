# coding=utf-8
import win32com.client as win32
import numpy as np
import pandas as pd
import myfun
import os, re, sys, time, datetime

### function generate email, then send or show based on parameter 'auto'
def CreateEmail(recipient, subject, msg, isattach, attachloc):

	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	mail.SentOnBehalfOfName = 'SHCOMIG@investorsgroup.com' #'Business Reporting (IG)'
	mail.To = recipient
	mail.Subject = subject
	mail.HTMLBody = msg
	if isattach == 1:
		mail.Attachments.Add(Source = attachloc)
		
	mail.Close(0)

	
template1 = '''On April 10, 2018 you received an email about your Federal tax withholding rate. 
This is a reminder that if you wish to change your Federal Tax rate when the new compensation administration system is implemented you must email <a href="mailto:SHCOMIG@investorsgroup.com">Compensation Mailbox - IG</a> by April 30, 2018. 
Pease confirm the rate of tax you wish to withhold in your email. Until that time your rates will remain set as they currently are.<br/><br/>
If you have any questions please email <a href="mailto:SHCOMIG@investorsgroup.com">Compensation Mailbox - IG</a>.
'''		

template2 = u'''Le 10 avril 2018, vous avez reçu un courriel au sujet de votre taux de retenue d’impôt fédéral. 
Nous désirons vous rappeler que si vous souhaitez modifier votre taux d’imposition fédéral lorsque le nouveau système d’administration de la rémunération sera en place, vous devez envoyer un courriel à la boîte <a href="mailto:SHCOMIG@investorsgroup.com">Compensation Mailbox – IG</a> d’ici le 30 avril 2018. 
Veuillez confirmer le taux de retenue souhaité dans votre courriel. D’ici là, vos taux actuels continueront de s’appliquer.<br/><br/>
Si vous avez des questions, écrivez à la boîte <a href="mailto:SHCOMIG@investorsgroup.com">Compensation Mailbox – IG</a>.
'''		

rlist1 = []
sublist1 = []
msglist1 = []
isattach1 = []
filelist1 = []

df = pd.read_excel('''C:\\pycode\\Data\\WEST2 Copy of 2018 0402 Federal Tax Withholding for Emails.xlsx''', sheetname=0)

for index, row in df.iterrows():
	if row['LANG'] == 'English':
		print 'process 1'
		rlist1.append(row['Email'])
		#rlist1.append('west.wang@investorsgroup.com')
		#sublist1.append(u'''PIP Company Account Tax Reporting''')
		sublist1.append(u'''Federal Tax Withholding - Reminder''')
		msglist1.append(template1)
		isattach1.append(0)
		filelist1.append('')
	else:
		rlist1.append(row['Email'])
		#rlist1.append('west.wang@investorsgroup.com')
		#sublist1.append(u'''PIP Company Account Tax Reporting''')
		sublist1.append(u'''Impôt fédéral prélevé à la source - Rappel''')
		msglist1.append(template2)
		isattach1.append(0)
		filelist1.append('')	
	
for e, s, m, a, f in zip(rlist1, sublist1, msglist1, isattach1, filelist1):
	CreateEmail(e, s, m, a, f)
