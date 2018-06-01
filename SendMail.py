# coding=utf-8
#---------------------------------------
#Read Outlook draft folder, then send all emails under draft folder
#---------------------------------------
import win32com.client as win32
import numpy as np
import pandas as pd

#read draft folder
outlook = win32.Dispatch('outlook.application')
namespace = outlook.GetNamespace("MAPI")
folders = namespace.Folders
dfolders = folders("West.Wang@investorsgroup.com").Folders("Drafts") #need to update outlook id

print 'total emails to be sent are'
print dfolders.Items.Count
#print dfolders.Items.Item(dfolders.Items.Count).To

#send all emails
numofmail = dfolders.Items.Count
for i in range(0, numofmail):
	print numofmail - i
	dfolders.Items.Item(numofmail - i).Send()

