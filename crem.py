import win32com.client as win32
import numpy as np
import pandas as pd

### function generate email, then send or show based on parameter 'auto'
def Emailer(recipient, subject, msg, auto=True):

	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	#mail.Sender = 'shrepan@investorsgroup.com'
	mail.SentOnBehalfOfName = 'Business Reporting (IG)'
	mail.To = recipient
	mail.Subject = subject
	mail.HTMLBody = msg
	#print 'in emailer'
	#mail.Body = msg
	if auto:
		mail.Send
	else:
		mail.Display(True) #show email

### template for Original allocation to PIP or SOP only		
template1 = '''Sent on behalf of Corinne Fontaine, Manager, Consultant Compensation & Reporting<br/><br/>
Great-West Life has recently advised us of an error in the medical/dental premiums that they sent us for our use in the Benefit Credit Re-enrollment 	process.   They inadvertently included Manitoba sales tax of 8% in the premium amount.   As a result, your medical/dental premiums on the Benefit Credit function in Pathway are overstated by 8% which resulted in an understatement of your available benefit credits for allocation to PIP and/or SOP by the 8% sales tax amount. <br/><br/>
You have {a} in available unallocated annual benefit credits.  Since you allocated all of your benefit credits to {b}, we will allocate this amount as well  to {b} which will result in an increase to your monthly personal matching deduction equal to twice the above amount spread over 11 months (October 2017 to August 2018).<br/><br/>
<b>If you wish to forfeit these additional unallocated benefit credits please advise us by Friday October 6, 2017</b> by sending an email to <a href="mailto:Charmaine.Falk@investorsgroup.com">Charmaine Falk</a>.  Otherwise we will proceed with allocating the additional credits. <br/><br/>
Please accept our sincere apologies for this error.'''		
		
template21 = '''Sent on behalf of Corinne Fontaine, Manager, Consultant Compensation & Reporting<br/><br/>
Great-West Life has recently advised us of an error in the medical/dental premiums that they sent us for our use in the Benefit Credit Re-enrollment 	process.   They inadvertently included Manitoba sales tax of 8% in the premium amount.   As a result, your medical/dental premiums on the Benefit Credit function in Pathway are overstated by 8% which resulted in an understatement of your available benefit credits for allocation to PIP and/or SOP by the 8% sales tax amount. <br/><br/>
You have {a} in available annual benefit credits.  Since you allocated all of your benefit credits to both PIP and SOP, we will allocate this amount as well to {b} which will result in an increase to your monthly premium equal to twice the above amount spread over 11 months (October 2017 to August 2018). <br/><br/>
<b>If you wish to forfeit these additional unallocated benefit credits please advise us by Friday October 6, 2017</b> by sending an email to <a href="mailto:Charmaine.Falk@investorsgroup.com">Charmaine Falk</a>.  Otherwise we will proceed with allocating the additional credits. <br/><br/>
Please accept our sincere apologies for this error.'''

template22 = '''Sent on behalf of Corinne Fontaine, Manager, Consultant Compensation & Reporting<br/><br/>
Great-West Life has recently advised us of an error in the medical/dental premiums that they sent us for our use in the Benefit Credit Re-enrollment process.   They inadvertently included Manitoba sales tax of 8% in the premium amount.   As a result, your medical/dental premiums on the Benefit Credit function in Pathway are overstated by 8% which resulted in an understatement of your available benefit credits for allocation to PIP and/or SOP by the 8% sales tax amount.  <br/><br/>
You have {a} in available annual benefit credits.  Since you allocated all of your benefit credits to both PIP and SOP, we will allocate {b} to {c} and {d} to {e} which will result in an increase to your monthly premium equal to twice the above amount spread over 11 months (October 2017 to August 2018). <br/><br/>
<b>If you wish to forfeit these additional unallocated benefit credits please advise us by Friday October 6, 2017</b> by sending an email to <a href="mailto:Charmaine.Falk@investorsgroup.com">Charmaine Falk</a>.  Otherwise we will proceed with allocating the additional credits. <br/><br/>
Please accept our sincere apologies for this error.'''

#### handle Original allocation to PIP or SOP only scenario
#rlist1 = []
#sublist1 = []
#msglist1 = []
#
#df = pd.read_excel('''Gail's Copy Sept 21 2017 Master Sept 6 2017.xlsx''', sheetname='Emails for PIP')
#df = df.append(pd.read_excel('''Gail's Copy Sept 21 2017 Master Sept 6 2017.xlsx''', sheetname='Emails for SOP'), ignore_index=True)
##print df
#
#for index, row in df.iterrows():
#		rlist1.append(row['Email'])
#		#rlist1.append('West.Wang@investorsgroup.com')
#		sublist1.append('Benefit Credits Adjustment')
#		#sublist1.append(row['Email'])
#		msglist1.append(template1.format(a = row['Amount of Benefit Credits to Allocate'], b = row['ALLOCATE TO:']))
#	
#for e, s, m in zip(rlist1, sublist1, msglist1):
#	Emailer(e, s, m, False)
#	
### handle Original allocation to PIP and/or SOP scenario
rlist2 = []
sublist2 = []
msglist2 = []	

df2 = pd.read_excel('''Gail's Copy Sept 21 2017 Master Sept 6 2017.xlsx''', sheetname='Emails for PIP and SOP')	

for index, row in df2.iterrows():
	if pd.notnull(row['Amount of Benefit Credits to Allocate.1']):
		print row['Amount of Benefit Credits to Allocate.1']	
		rlist2.append(row['Email'])
		#rlist2.append('West.Wang@investorsgroup.com')
		sublist2.append('Benefit Credits Adjustment')	
		#sublist2.append(row['Email'])
		msglist2.append(template22.format(a=row['Total Amount to allocate'], b=row['Amount of Benefit Credits to Allocate'], c=row['ALLOCATE TO:'], d=row['Amount of Benefit Credits to Allocate.1'], e=row['ALLOCATE TO:.1']))
	else:
		rlist2.append(row['Email'])
		#rlist2.append('West.Wang@investorsgroup.com')
		sublist2.append('Benefit Credits Adjustment')
		#sublist2.append(row['Email'])
		msglist2.append(template21.format(a = row['Amount of Benefit Credits to Allocate'], b = row['ALLOCATE TO:']))		
		
for e, s, m in zip(rlist2, sublist2, msglist2):
	Emailer(e, s, m, False)
