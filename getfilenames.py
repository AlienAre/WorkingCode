import os, re, sys, time, xlrd, pyodbc, datetime
import pandas as pd
import csv

filelist = [] 

for file in os.listdir('C:\\PDF Files'):
	if file.endswith('.pdf'):
		filelist.append(file)
print filelist

with open('filename.csv', 'wb') as myfile:
    wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
    wr.writerow(filelist)