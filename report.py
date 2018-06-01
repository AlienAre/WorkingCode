import numpy as np
import pandas as pd

#df = pd.read_excel('F:\\Files For\\Hai Yen Nguyen\\RO Reserve Fund\\2017-12\\RO.xlsx', sheet_name='CsltBal')
#
#del df['LKG_CSLT_SMPL_DTE']
#del df['RO No']
#
#df = pd.pivot_table(df,index=['RO','RD No', 'RD Name', 'Status', 'Cslt No', 'Start Date'],values=['Comm Account Bal', 'AP Bal', 'Total Debt'],aggfunc=np.sum, margins=True)
#
## Create a Pandas Excel writer using XlsxWriter as the engine.
#writer = pd.ExcelWriter('rep.xlsx', engine='xlsxwriter')
#
## Convert the dataframe to an XlsxWriter Excel object.
#df.to_excel(writer, index=True)
#
## Close the Pandas Excel writer and output the Excel file.
#writer.save()
df = pd.read_excel('C:\pycode\\RO.xlsx', sheet_name='CsltBal')

print df['RO'].unique()
print df[df['RO'] == 9]
writer = pd.ExcelWriter('output.xlsx')

for i in df['RO'].unique():
    temp_df = df[df['RO'] == i]
    temp_df.to_excel(writer,'RO' + str(i))

writer.save()