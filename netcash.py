import re, sys, time
import pandas as pd
import calendar as cr
import datetime
from datetime import date
import pyodbc
import myfun

#accept user input to decide which month they want to run
reportenddate = datetime.datetime.strptime(raw_input('Please enter the month end date you want to process \(date format: mm/dd/yyyy\):'), '%m/%d/%Y')
start_time = datetime.datetime.now()
#enddate = datetime.date(int(time.strftime("%Y")), int(time.strftime("%m"))-1, cr.monthrange(int(time.strftime("%Y")), int(time.strftime("%m"))-1)[1]).strftime("%d/%m/%Y")
#print enddate

#---------------------------------------------
startday = myfun.getCycleStartDate(datetime.datetime.strptime('01/01/2018', '%m/%d/%Y'))
endday = myfun.getCycleEndDate(datetime.datetime.strptime('01/01/2018', '%m/%d/%Y'))
#lastmonthend = myfun.getLastMonthEndDate(date.today())
lastmonthend = myfun.getLastMonthEndDate(reportenddate)
lastquarterenddate = myfun.getLastQuarterEndDate(reportenddate)
print 'startdate is ', startday
print 'endday is ', endday
print 'lastmonthend is ', lastmonthend
#print 'last2monthend is ', last2monthend
print 'lastquarterenddate is ', lastquarterenddate
#sys.exit("this is end")
#---------------------------------------------

numofperiodleft = 0

driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"
db_file = r"F:\\3-Compensation Programs\\Net Flows Bonus\\NetFlow.accdb;"
user = "admin"
password = ""
odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)
conn = pyodbc.connect(odbc_conn_str)

sql = ''' SELECT COUNT(D.CycleEndDate) FROM DimCycleDate D WHERE D.Year = ''' + str(endday.year) + ''' AND D.Quarter = ''' + str(myfun.getQuarter(endday)) + ''' AND D.CycleEndDate >= # ''' + str(endday) + ''' # '''

cur = conn.cursor()
cur.execute(sql)
numofperiodleft = cur.fetchone()[0]
print numofperiodleft

######## Timer ######
elapsed = datetime.datetime.now() - start_time
start_time = datetime.datetime.now()
print datetime.datetime.now()
print 'Get quarter info finished, exec time is ' + str(elapsed)


###### Get openning assets prorated before current cycle
sql = '''
		SELECT C.LKG_CSLT_RO_NUM AS RO
			,SUM(IIF(ISNULL(A.AUM), 0, A.AUM/6)) AS AMT
		FROM BRANUSER_BRAN_LKG_CSLT AS C
		LEFT JOIN (
			SELECT tCslt_Scorecard.Cslt AS Cslt
				,tCslt_Scorecard.Assets_MF_GIF_AUM AS AUM
			FROM tCslt_Scorecard
			WHERE tCslt_Scorecard.Smple_Dte = # ''' + str(lastquarterenddate) + ''' #
			) AS A ON A.Cslt = C.LKG_CSLT_NUM
		WHERE (
				C.LKG_CSLT_SMPL_DTE IN (
					SELECT D.CycleEndDate
					FROM DimCycleDate D
					WHERE D.Year = ''' + str(endday.year) + '''
						AND D.Quarter = ''' + str(myfun.getQuarter(endday)) + '''
						AND D.CycleEndDate < # ''' + str(endday) + ''' #
					)
				)
		GROUP BY C.LKG_CSLT_RO_NUM
		'''
#print sql
openassetdf = pd.read_sql_query(sql,conn)

######## Timer ######
elapsed = datetime.datetime.now() - start_time
start_time = datetime.datetime.now()
print datetime.datetime.now()
print 'Get openning assets prorated before current cycle finished, exec time is ' + str(elapsed)


#### Get openning assets prorated for the rest of the quarter
sql = '''
		SELECT C.LKG_CSLT_RO_NUM AS RO
			,''' + str(numofperiodleft) + ''' * SUM(IIF(ISNULL(A.AUM), 0, A.AUM/6)) AS AMT
		FROM BRANUSER_BRAN_LKG_CSLT AS C
		LEFT JOIN (
			SELECT tCslt_Scorecard.Cslt AS Cslt
				,tCslt_Scorecard.Assets_MF_GIF_AUM AS AUM
			FROM tCslt_Scorecard
			WHERE tCslt_Scorecard.Smple_Dte = # ''' + str(lastquarterenddate) + ''' #
			) AS A ON A.Cslt = C.LKG_CSLT_NUM
		WHERE (
				C.LKG_CSLT_SMPL_DTE IN (
					SELECT D.CycleEndDate
					FROM DimCycleDate D
					WHERE D.Year = ''' + str(endday.year) + '''
						AND D.Quarter = ''' + str(myfun.getQuarter(endday)) + '''
						AND D.CycleEndDate = # ''' + str(endday) + ''' #
					)
				)
		GROUP BY C.LKG_CSLT_RO_NUM
		'''
#print sql
openassetdf = openassetdf.append(pd.read_sql_query(sql,conn), ignore_index=True)
######## Timer ######
elapsed = datetime.datetime.now() - start_time
start_time = datetime.datetime.now()
print datetime.datetime.now()
print 'Get rest of openning assets finished, exec time is ' + str(elapsed)

openassetalldf = openassetdf.groupby(['RO'], as_index=False)['AMT'].sum()

######## Timer ######
elapsed = datetime.datetime.now() - start_time
start_time = datetime.datetime.now()
print datetime.datetime.now()
print 'Get all openning assets finished, exec time is ' + str(elapsed)

#### Get Monthly New Business amount by RO
sql = '''
		SELECT RO_Num AS RO, SUM(CurrYr_NNB) AS NB FROM
		(
			SELECT RO_Num, CurrYr_NNB FROM tRO_Scorecard WHERE Date = # ''' + str(reportenddate) + ''' #
			UNION ALL
			SELECT RO_Num, -1 * CurrYr_NNB FROM tRO_Scorecard WHERE Date = #''' + str(lastmonthend) + ''' #
		) AS A
		GROUP BY
			RO_Num
		'''
#print sql
newbusdf = pd.read_sql_query(sql,conn)

######## Timer ######
elapsed = datetime.datetime.now() - start_time
start_time = datetime.datetime.now()
print datetime.datetime.now()
print 'Get Monthly New Business finished, exec time is ' + str(elapsed)

cur.close()
conn.close()

outputdf = openassetalldf.merge(newbusdf, how='left', on='RO')

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('C:\\pycode\\netcash.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
#openassetdf.to_excel(writer, sheet_name='oa', index=False)
#openassetalldf.to_excel(writer, sheet_name='oaall', index=False)
outputdf.to_excel(writer, sheet_name='netflow', index=False)
# Close the Pandas Excel writer and output the Excel file.
writer.save()

######## Timer ######
elapsed = datetime.datetime.now() - start_time
start_time = datetime.datetime.now()
print datetime.datetime.now()
print 'All job finished, exec time is ' + str(elapsed)
