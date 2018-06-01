# coding=utf-8
import win32com.client as win32
import numpy as np
import pandas as pd
import myfun
import os, re, sys, time, datetime

### function generate email, then send or show based on parameter 'auto'
def CreateEmail(recipient, subject, msg, isattach, attachloc):
	print 'Generate'
	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	mail.SentOnBehalfOfName = 'Business Reporting (IG)' #'shcomig@investorsgroup.com'
	mail.To = recipient
	mail.Bcc = 'Glenda.Dulatre@investorsgroup.com'
	mail.Subject = subject
	mail.HTMLBody = msg
	if isattach == 1:
		mail.Attachments.Add(Source = attachloc)
		
	mail.Close(0)

		
### template for Original allocation to PIP or SOP only		
template1 = u'''As announced in the recent AdvantagePlus news item <b>“Your compensation statement is going digital”</b> the compensation administration system is being redeveloped.  As a result, the way in which Federal Tax is withheld is also changing.
Today you can set a different tax withholding rate for each of your regular income and your asset retention bonus income.  
In the new system there is only one federal tax withholding rate that will apply to all of your commission and asset income.<br/><br/>
You currently have set a federal tax rate of 0% for your regular income and a federal tax rate of {y}% for your asset retention bonus (ARB) income.  
When the new system is implemented, since your Federal Tax rate is set at 0%, you will not have any tax withheld on any income.<br/><br/>
Please email <a href="mailto:shcomig@investorsgroup.com">Compensation Mailbox - IG</a> by April 30, 2018 if you wish to change your Federal Tax rate from 0% when the new system is implemented.  
Please confirm the rate of tax you wish to withhold in your email.   Until that time your rate will remain set as it currently is.<br/><br/>
If you have any questions please email <a href="mailto:shcomig@investorsgroup.com">Compensation Mailbox –IG</a>.  
'''	
template2 = u'''As announced in the recent AdvantagePlus news item <b>“Your compensation statement is going digital”</b> the compensation administration system is being redeveloped.  As a result, the way in which Federal Tax is withheld is also changing.
Today you can set a different tax withholding rate for each of your regular income and your asset retention bonus income.  
In the new system there is only one federal tax withholding rate that will apply to all of your commission and asset income.<br/><br/>
You currently have a federal tax rate of {x}% set for your regular income and a federal tax rate of {y}% set for your asset retention bonus income.   
When the new system is implemented the tax rate you have set for your regular income of {x}% will also apply to all of your asset compensation.<br/><br/>
Please email <a href="mailto:shcomig@investorsgroup.com">Compensation Mailbox - IG</a> by April 30, 2018 if you wish to change your Federal Tax rate from 0% when the new system is implemented.  
Please confirm the rate of tax you wish to withhold in your email.   Until that time your rate will remain set as it currently is.<br/><br/>
If you have any questions please email <a href="mailto:shcomig@investorsgroup.com">Compensation Mailbox –IG</a>.  
'''	
template3 = u'''As announced in the recent AdvantagePlus news item <b>“Your compensation statement is going digital”</b> the compensation administration system is being redeveloped.  As a result, the way in which Federal Tax is withheld is also changing.
Today you can set a different tax withholding rate for each of your regular income and your asset retention bonus income.  
In the new system there is only one federal tax withholding rate that will apply to all of your commission and asset income.<br/><br/>
You currently have set a federal tax rate of {x}% for your regular income and a federal tax rate of 0% on your asset retention bonus income.  
When the new system is implemented the tax rate you have set for your regular income of {x}% will also apply to all of your asset compensation.<br/><br/>
Please email <a href="mailto:shcomig@investorsgroup.com">Compensation Mailbox - IG</a> by April 30, 2018 if you wish to change your Federal Tax rate from 0% when the new system is implemented.  
Please confirm the rate of tax you wish to withhold in your email.   Until that time your rate will remain set as it currently is.<br/><br/>
If you have any questions please email <a href="mailto:shcomig@investorsgroup.com">Compensation Mailbox –IG</a>.  
'''	
template4 = u'''As announced in the recent AdvantagePlus news item <b>“Your compensation statement is going digital”</b> the compensation administration system is being redeveloped.  As a result, the way in which Federal Tax is withheld is also changing.
Today you can set a different tax withholding rate for each of your regular income and your asset retention bonus income.  
In the new system there is only one federal tax withholding rate that will apply to all of your commission and asset income.<br/><br/>
You currently have set a federal tax rate of {y}% for your asset retention bonus income and you do not have a federal tax withholding rate set for your regular income.  
When the new system is implemented since you do not have a Federal Tax rate set on regular income, you will not have any tax withheld on any income.<br/><br/>
Please email <a href="mailto:shcomig@investorsgroup.com">Compensation Mailbox - IG</a> by April 30, 2018 if you wish to change your Federal Tax rate from 0% when the new system is implemented.  
Please confirm the rate of tax you wish to withhold in your email.   Until that time your rate will remain set as it currently is.<br/><br/>
If you have any questions please email <a href="mailto:shcomig@investorsgroup.com">Compensation Mailbox –IG</a>.  
'''		
templateF1 = u'''Comme il a été annoncé dans l’article <b>Votre relevé de rémunération passe au numérique!</b> publié récemment à la section Actualité d’AvantagePlus, nous revoyons actuellement notre système d’administration de la rémunération, ce qui entraînera des changements dans la façon dont l’impôt fédéral est prélevé à la source. Présentement, vous pouvez établir un taux de retenue d’impôt distinct pour votre revenu régulier et votre prime de rétention de l’actif. 
Dans le nouveau système, un taux de retenue d’impôt fédéral unique s’appliquera à vos commissions et à votre revenu tiré de l’actif.<br/><br/>

Votre taux d’imposition fédéral est présentement établi à 0% pour votre revenu régulier et à {y}% pour votre prime de rétention de l’actif (PRA). Quand le nouveau système sera en place, puisque votre taux d’imposition fédéral est de 0%, 
aucune retenue d’impôt ne sera opérée sur vos revenus.<br/><br/>
Écrivez à la boîte courriel <a href="mailto:shcomig@investorsgroup.com">Compensation Mailbox - IG</a> d’ici le 30 avril 2018 si vous souhaitez que des retenues d’impôt fédéral soient effectuées sur vos revenus quand le nouveau système sera en place. Veuillez confirmer le taux de retenue souhaité dans votre courriel. D’ici là, vos taux actuels continueront de s’appliquer.<br/><br/>
Si vous avez des questions, écrivez à la boîte <a href="mailto:shcomig@investorsgroup.com">Compensation Mailbox – IG</a>.  
'''	
templateF2 = u'''Comme il a été annoncé dans l’article <b>Votre relevé de rémunération passe au numérique!</b> publié récemment à la section Actualité d’AvantagePlus, nous revoyons actuellement notre système d’administration de la rémunération, ce qui entraînera des changements dans la façon dont l’impôt fédéral est prélevé à la source. Présentement, vous pouvez établir un taux de retenue d’impôt distinct pour votre revenu régulier et votre prime de rétention de l’actif. 
Dans le nouveau système, un taux de retenue d’impôt fédéral unique s’appliquera à vos commissions et à votre revenu tiré de l’actif.<br/><br/>

Votre taux d’imposition fédéral est présentement établi à {x}% pour votre revenu régulier et à {y}% pour votre prime de rétention de l’actif (PRA). Quand le nouveau système sera en place, le taux de retenue de {x}% que vous avez établi pour votre revenu régulier s’appliquera aussi à votre revenu tiré de l’actif.<br/><br/>
Écrivez à la boîte courriel <a href="mailto:shcomig@investorsgroup.com">Compensation Mailbox - IG</a> d’ici le 30 avril 2018 si vous souhaitez que des retenues d’impôt fédéral soient effectuées sur vos revenus quand le nouveau système sera en place. Veuillez confirmer le taux de retenue souhaité dans votre courriel. D’ici là, vos taux actuels continueront de s’appliquer.<br/><br/>
Si vous avez des questions, écrivez à la boîte <a href="mailto:shcomig@investorsgroup.com">Compensation Mailbox – IG</a>.  
'''	
templateF3 = u'''Comme il a été annoncé dans l’article <b>Votre relevé de rémunération passe au numérique!</b> publié récemment à la section Actualité d’AvantagePlus, nous revoyons actuellement notre système d’administration de la rémunération, ce qui entraînera des changements dans la façon dont l’impôt fédéral est prélevé à la source. Présentement, vous pouvez établir un taux de retenue d’impôt distinct pour votre revenu régulier et votre prime de rétention de l’actif. 
Dans le nouveau système, un taux de retenue d’impôt fédéral unique s’appliquera à vos commissions et à votre revenu tiré de l’actif.<br/><br/>

Votre taux d’imposition fédéral est présentement établi à {x}% pour votre revenu régulier et à 0% pour votre prime de rétention de l’actif (PRA). Quand le nouveau système sera en place, le taux de retenue de {x}% que vous avez établi pour votre revenu régulier s’appliquera aussi à votre revenu tiré de l’actif.<br/><br/>
Écrivez à la boîte courriel <a href="mailto:shcomig@investorsgroup.com">Compensation Mailbox - IG</a> d’ici le 30 avril 2018 si vous souhaitez que des retenues d’impôt fédéral soient effectuées sur vos revenus quand le nouveau système sera en place. Veuillez confirmer le taux de retenue souhaité dans votre courriel. D’ici là, vos taux actuels continueront de s’appliquer.<br/><br/>
Si vous avez des questions, écrivez à la boîte <a href="mailto:shcomig@investorsgroup.com">Compensation Mailbox – IG</a>.  
'''	
templateF4 = u'''Comme il a été annoncé dans l’article <b>Votre relevé de rémunération passe au numérique!</b> publié récemment à la section Actualité d’AvantagePlus, nous revoyons actuellement notre système d’administration de la rémunération, ce qui entraînera des changements dans la façon dont l’impôt fédéral est prélevé à la source. Présentement, vous pouvez établir un taux de retenue d’impôt distinct pour votre revenu régulier et votre prime de rétention de l’actif. 
Dans le nouveau système, un taux de retenue d’impôt fédéral unique s’appliquera à vos commissions et à votre revenu tiré de l’actif.<br/><br/>

Le taux de l’impôt prélevé à la source sur votre prime de rétention de l’actif est présentement établi à {y}% et aucun impôt n’est prélevé à la source sur votre revenu régulier. Quand le nouveau système sera en place, comme aucun taux de retenue d’impôt fédéral n’est établi sur votre revenu régulier, aucun impôt ne sera prélevé à la source sur l’ensemble de vos revenus.<br/><br/>
Écrivez à la boîte courriel <a href="mailto:shcomig@investorsgroup.com">Compensation Mailbox - IG</a> d’ici le 30 avril 2018 si vous souhaitez que des retenues d’impôt fédéral soient effectuées sur vos revenus quand le nouveau système sera en place. Veuillez confirmer le taux de retenue souhaité dans votre courriel. D’ici là, vos taux actuels continueront de s’appliquer.<br/><br/>
Si vous avez des questions, écrivez à la boîte <a href="mailto:shcomig@investorsgroup.com">Compensation Mailbox – IG</a>.  
'''	

template = u'''Message envoyé au nom de Gail Purcell, directrice générale, Rémunération des conseillers et information d’affaires<br/><br/>
Votre relevé de compte du PPC en date du 31 décembre indique sur un feuillet T5008 le gain en capital réalisé à l’acquisition des cotisations de la société dans le Programme de placement contributif (PPC). Ce gain en capital sera également déclaré sur un feuillet T3. Veuillez ne pas tenir compte du feuillet T5008 pour le compte PPC de la société et déclarer le gain en capital du T3 dans votre déclaration de revenus.
<br/><br/>
Si vous avez des questions, veuillez écrire à la boîte courriel <a href="mailto:shrepan@investorsgroup.com">Business Reporting (IG)</a>.
'''		

rlist1 = []
sublist1 = []
msglist1 = []
isattach1 = []
filelist1 = []

df = pd.read_excel('''TaxEmails.xlsx''', sheetname=0)
#print df.head()
#print df.dtypes
#for ro in df['RO']:
#	print myfun.addZero(ro.astype(str), 3)
#sys.exit('Process is stopped')

for index, row in df.iterrows():
	if row['LANGUAGE'] == 'English':
		#print 'process Eng'
		#print row['Letter']
		if str(row['Letter']) == '1':
			print 'process 1'
			rlist1.append(row['EMAIL'])
			#rlist1.append('west.wang@investorsgroup.com')
			#sublist1.append(u'''PIP Company Account Tax Reporting''')
			sublist1.append(u'''Federal Tax Withholding''')
			msglist1.append(template1.format(y=row['y']))
			isattach1.append(0)
			filelist1.append('')
		elif str(row['Letter']) == '2':
			print 'process 2'
			rlist1.append(row['EMAIL'])
			#rlist1.append('west.wang@investorsgroup.com')
			#sublist1.append(u'''PIP Company Account Tax Reporting''')
			sublist1.append(u'''Federal Tax Withholding''')
			msglist1.append(template2.format(x=row['x'], y=row['y']))
			isattach1.append(0)
			filelist1.append('')			
		elif str(row['Letter']) == '3':
			rlist1.append(row['EMAIL'])
			#rlist1.append('west.wang@investorsgroup.com')
			#sublist1.append(u'''PIP Company Account Tax Reporting''')
			sublist1.append(u'''Federal Tax Withholding''')
			msglist1.append(template3.format(x=row['x']))
			isattach1.append(0)
			filelist1.append('')			
		elif str(row['Letter']) == '4':
			rlist1.append(row['EMAIL'])
			#rlist1.append('west.wang@investorsgroup.com')
			#sublist1.append(u'''PIP Company Account Tax Reporting''')
			sublist1.append(u'''Federal Tax Withholding''')
			msglist1.append(template4.format(y=row['y']))
			isattach1.append(0)
			filelist1.append('')			
	else:
		if str(row['Letter']) == '1':
			rlist1.append(row['EMAIL'])
			#rlist1.append('west.wang@investorsgroup.com')
			#sublist1.append(u'''PIP Company Account Tax Reporting''')
			sublist1.append(u'''Impôt fédéral prélevé à la source''')
			msglist1.append(templateF1.format(y=row['y']))
			isattach1.append(0)
			filelist1.append('')
		elif str(row['Letter']) == '2':
			rlist1.append(row['EMAIL'])
			#rlist1.append('west.wang@investorsgroup.com')
			#sublist1.append(u'''PIP Company Account Tax Reporting''')
			sublist1.append(u'''Impôt fédéral prélevé à la source''')
			msglist1.append(templateF2.format(x=row['x'], y=row['y']))
			isattach1.append(0)
			filelist1.append('')			
		elif str(row['Letter']) == '3':
			rlist1.append(row['EMAIL'])
			#rlist1.append('west.wang@investorsgroup.com')
			#sublist1.append(u'''PIP Company Account Tax Reporting''')
			sublist1.append(u'''Impôt fédéral prélevé à la source''')
			msglist1.append(templateF3.format(x=row['x']))
			isattach1.append(0)
			filelist1.append('')			
		elif str(row['Letter']) == '4':
			rlist1.append(row['EMAIL'])
			#rlist1.append('west.wang@investorsgroup.com')
			#sublist1.append(u'''PIP Company Account Tax Reporting''')
			sublist1.append(u'''Impôt fédéral prélevé à la source''')
			msglist1.append(templateF4.format(y=row['y']))
			isattach1.append(0)
			filelist1.append('')			
#	rlist1.append(row['Email'])
#	subject = u'''2017 Consultant Debt Report RO: ''' + myfun.addZero(str(row['RO']), 3)
#	sublist1.append(subject)
#	msglist1.append(template1)
#	isattach1.append(1)
#	file = 'F:\\4-Compensation Reporting\\4-Consultant Financing\\1-Cslt Debt Reporting\\2017\\201712\\Region Office ' + myfun.addZero(str(row['RO']), 3) + ' Debt Report 2017-12.pdf'
#	filelist1.append(file)

for e, s, m, a, f in zip(rlist1, sublist1, msglist1, isattach1, filelist1):
	CreateEmail(e, s, m, a, f)
