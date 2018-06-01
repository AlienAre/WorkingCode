import os, re, sys, time, xlrd, pyodbc, datetime
from datetime import date
import fnmatch
import numpy as np
import pandas as pd
import itertools as it
from openpyxl import load_workbook
import myfun as dd

df01 = pd.read_csv('C:\\Users\\wangwe5\\Documents\\MF ASF\\201803\\asf201802.txt', sep='|')
df02 = pd.read_csv('C:\\Users\\wangwe5\\Documents\\MF ASF\\201803\\asf201803.txt', sep='|')

df01['MFsum1'] = df01.iloc[:,40] + df01.iloc[:,41]
df01['GIFsum1'] = df01.iloc[:,42] + df01.iloc[:,43]
df01total = df01.iloc[:,[1,46,47]]
df01total = df01total.groupby(['CONSULTANT NUMBER                            '], as_index=False)['MFsum1','GIFsum1'].sum()

df02['MFsum2'] = df02.iloc[:,40] + df02.iloc[:,41]
df02['GIFsum2'] = df02.iloc[:,42] + df02.iloc[:,43]
df02total = df02.iloc[:,[1,46,47]]
df02total = df02total.groupby(['CONSULTANT NUMBER                            '], as_index=False)['MFsum2','GIFsum2'].sum()

dfasf = pd.read_csv('C:\\Users\\wangwe5\\Documents\\MF ASF\\201803\\ASF04302018.txt', sep='|')
#print dfasf.dtypes
dfasf['MF NL QA AVG'] = dfasf['MF NL QA AVG'].str.replace(',','').astype(np.float64)
dfasf['MF ASF AMT'] = dfasf['MF ASF AMT'].str.replace(',','').astype(np.float64)
dfasf['GIF NL QA AVG'] = dfasf['GIF NL QA AVG'].str.replace(',','').astype(np.float64)
dfasf['GIF ASF AMT'] = dfasf['GIF ASF AMT'].str.replace(',','').astype(np.float64)
dfasftotal = dfasf.groupby(['CYCLE DATE','REG OFF NUM','DIV OFF NUM','CNSLT NUM','LAST NAME','FIRST NAME','TERM IND','SPEC IND','PAID AL','ASF RATE'], as_index=False).agg({'COMM ACCT': max, 'MF NL QA AVG': sum,'MF ASF AMT': sum,'GIF NL QA AVG': sum,'GIF ASF AMT': sum})
dfasftotal = dfasftotal[['CYCLE DATE','REG OFF NUM','DIV OFF NUM','CNSLT NUM','COMM ACCT','LAST NAME','FIRST NAME','TERM IND','SPEC IND','PAID AL','ASF RATE','MF NL QA AVG','MF ASF AMT','GIF NL QA AVG','GIF ASF AMT']]
dfasftotal.sort_values(['CNSLT NUM'])

dfoutput = pd.merge(dfasftotal, df01total, how='left', left_on='CNSLT NUM', right_on='CONSULTANT NUMBER                            ')
dfoutput = pd.merge(dfoutput, df02total, how='left', left_on='CNSLT NUM', right_on='CONSULTANT NUMBER                            ')
dfoutput['MFsum1'].fillna(0, inplace = True)
dfoutput['GIFsum1'].fillna(0, inplace = True)
dfoutput['MFsum2'].fillna(0, inplace = True)
dfoutput['GIFsum2'].fillna(0, inplace = True)
dfoutput['MF NL QA AVG'] = dfoutput[['MFsum1', 'MFsum2']].mean(axis=1)
dfoutput['GIF NL QA AVG'] = dfoutput[['GIFsum1', 'GIFsum2']].mean(axis=1)

#dfoutput.to_csv('asf.csv')

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('asf.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
dfoutput.to_excel(writer, sheet_name='Sheet1', index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()