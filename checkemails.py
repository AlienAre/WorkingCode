import win32com.client as win32
import numpy as np
import pandas as pd
import csv

columns = ['to']
#df = pd.DataFrame(columns=columns)
myl = []

outlook = win32.Dispatch('outlook.application')	
mynnamespace = outlook.GetNamespace('MAPI')
#drafts = mynnamespace.Folders('Mailbox - Wang, West').Folders('Drafts')
drafts = mynnamespace.Folders.Item('Mailbox - Wang, West').Folders.Item('Drafts').Items
print drafts.Count
for i in drafts:
	myl.append(i.To)

df = pd.DataFrame(data=myl, columns=columns)
df.to_csv('list.csv', index=False)