# coding=utf-8
import win32com.client as win32
import numpy as np
import pandas as pd

#### function generate email, then send or show based on parameter 'auto'
#def CreateEmail(recipient, subject, msg):
#
#	outlook = win32.Dispatch('outlook.application')
#	mail = outlook.CreateItem(0)
#	mail.SentOnBehalfOfName = 'Business Reporting (IG)'
#	mail.To = recipient
#	mail.Subject = subject
#	mail.HTMLBody = msg
#	mail.Close(0)

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.SentOnBehalfOfName = 'Business Reporting (IG)'
mail.To = 'west.wang@investorsgroup.com'
mail.Subject = 'This is a test for att'
mail.HTMLBody = 'This is good'
#mail.Attachments.Add(Source = "C:\\pycode\\Code\\Email.xlsx")
mail.Send()
#mail.Close(0)
	
		
#### template for Original allocation to PIP or SOP only		
#template1 = '''Sent on behalf of Gail Purcell, Director, Consultant Compensation & Reporting<br/><br/>
#Your December 31 PIP Statement reports on a T5008 the capital gain realized on the vesting of the company contributions in the Participating Investment Program (PIP).   This capital gain will also be reported on a T3.   Please ignore the T5008 for the Company PIP account and report the capital gain from the T3 on your income tax return.
#<br/><br/>
#If you have any questions please email <a href="mailto:shrepan@investorsgroup.com">Business Reporting (IG)</a>
#'''		
#
#template2 = u'''Message envoyé au nom de Gail Purcell, directrice générale, Rémunération des conseillers et information d’affaires<br/><br/>
#Votre relevé de compte du PPC en date du 31 décembre indique sur un feuillet T5008 le gain en capital réalisé à l’acquisition des cotisations de la société dans le Programme de placement contributif (PPC). Ce gain en capital sera également déclaré sur un feuillet T3. Veuillez ne pas tenir compte du feuillet T5008 pour le compte PPC de la société et déclarer le gain en capital du T3 dans votre déclaration de revenus.
#<br/><br/>
#Si vous avez des questions, veuillez écrire à la boîte courriel <a href="mailto:shrepan@investorsgroup.com">Business Reporting (IG)</a>.
#'''		
#
#rlist1 = []
#sublist1 = []
#msglist1 = []
#
#df = pd.read_excel('''email.xlsx''', sheetname=0)
##print df
#
#for index, row in df.iterrows():
#	if row['L'] == 'E':
#		#rlist1.append(row['Email'])
#		rlist1.append('west.wang@investorsgroup.com')
#		#sublist1.append(u'''PIP Company Account Tax Reporting''')
#		sublist1.append('This is E test' + row['Email'])
#		msglist1.append(template1)
#	else:
#		#rlist1.append(row['Email'])
#		rlist1.append('west.wang@investorsgroup.com')
#		#sublist1.append(u'''Renseignements fiscaux pour le compte du PPC de la société''')
#		sublist1.append('This is F Test' + row['Email'])
#		msglist1.append(template2)		
#	
#for e, s, m in zip(rlist1, sublist1, msglist1):
#	CreateEmail(e, s, m)