import win32com.client as win32
import numpy as np
import pandas as pd


################################		
el = ['West.Wang@investorsgroup.com','West.Wang@investorsgroup.com','West.Wang@investorsgroup.com','West.Wang@investorsgroup.com']
sl = ['this is test','this is test','this is test','this is test']
ml = ['heel','heel','heel','heel']
################################

### function generate email, then send or show based on parameter 'auto'
def Emailer(recipient, subject, msg, auto=True):

	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	mail.SentOnBehalfOfName = 'Business Reporting (IG)'
	mail.To = recipient
	mail.Subject = subject
	mail.HTMLBody = msg
	mail.Close(0)
	#
	#mynnamespace = outlook.GetNamespace('MAPI')
	##print mynnamespace.ExchangeMailboxServerName
	#floders = mynnamespace.Folders.Item('Mailbox - Business Reporting (IG)').Folders.Item('Drafts')
	#print floders.Items.Count
	#
	#floders = mynnamespace.Folders.Item('Mailbox - Wang, West').Folders.Item('Drafts')
	#print floders.Items.Count	

def SendEmail():
	outlook = win32.Dispatch('outlook.application')	
	mynnamespace = outlook.GetNamespace('MAPI')
	#drafts = mynnamespace.Folders('Mailbox - Wang, West').Folders('Drafts')
	drafts = mynnamespace.Folders.Item('Mailbox - Wang, West').Folders.Item('Drafts')
	print drafts.Items.Count
	#for i in drafts:
	#	i.Send
	#	print 'Sent'
	for idx in range(drafts.Items.Count):
		print idx
		drafts.Items(idx + 1).Send
		
		
for e, s, m in zip(el, sl, ml):
	Emailer(e, s, m, False)

SendEmail()	

df.append({'to': i.To, 'subject': i.Subject}, ignore_index=True)