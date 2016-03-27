import openpyxl as xl
import pandas as p
import os
import numpy as nm
from numbers import Number
import datetime as d

############################ DATES ############################
ThisWeek = raw_input('<FORMAT: m/d/yyyy> Input date for this weeks report: ')
LastWeek = raw_input('<FORMAT: m/d/yyyy> Input date for last weeks report: ')
#ThisWeek = '3/24/2016'
#LastWeek = '3/17/2016'

ThisWeekInt = []
for i in ThisWeek.split('/'):
	ThisWeekInt.append(int(i))
ThisWeekDT = d.datetime(ThisWeekInt[2], ThisWeekInt[0], ThisWeekInt[1])
ThisWeekDash = ThisWeek.replace('/', '-')

LastWeekInt = []
for i in LastWeek.split('/'):
	LastWeekInt.append(int(i))
LastWeekDT = d.datetime(LastWeekInt[2], LastWeekInt[0], LastWeekInt[1])
LastWeekDash = LastWeek.replace('/', '-')


############################ PATHS ############################
#MAC
#DataPath = '/Users/jamespardee/Desktop/JPY_Post/II. Exposure Breakdown SDHY Backup '+ThisWeekDash+'.xlsx'
#InDir = '/Users/jamespardee/Desktop/JPY_Post/'
#OutPath ='/Users/jamespardee/Desktop/JPY_Post/Weekly Reporting_'+ThisWeekDash+'_SDHY.xlsx'
#LastWeekPath ='/Users/jamespardee/Desktop/JPY_Post/Weekly Reporting_'+LastWeekDash+'_SDHY.xlsx'

#WIN LOCAL
#ThisWeekFile = 'Weekly Reporting ' + ThisWeekDash + ' IG.xlsx'
#LastWeekFile = 'Weekly Reporting ' + LastWeekDash + ' IG.xlsx'
#DataPath = 'C:/Users/jpardee/Desktop/JPY_Post/II. Exposure Breakdown IG Backup '+ThisWeekDash+'.xlsx'
#InDir = 'C:/Users/jpardee/Desktop/JPY_Post/'
#OutPath = 'C:/Users/jpardee/Desktop/JPY_Post/' + ThisWeekDash + ' IG/' + ThisWeekFile
#LastWeekPath = 'C:/Users/jpardee/Desktop/JPY_Post/' + LastWeekDash + ' IG/' + LastWeekFile

#WIN DRIVE
ThisWeekFile = 'Weekly Reporting ' + ThisWeekDash + ' IG.xlsx'
LastWeekFile = 'Weekly Reporting ' + LastWeekDash + ' IG.xlsx'
DataPath = 'Y:/Non-US/NB East Asia (Japan & Korea)/Japan Post Bank/IG Credit Portfolio/Weekly/2016/Backup/II. Exposure Breakdown IG Backup '+ThisWeekDash+'.xlsx'
#InDir = 'C:\Users\\jpardee\\Desktop\\JPY_Post\\'
OutPath = 'Y:/Non-US/NB East Asia (Japan & Korea)/Japan Post Bank/IG Credit Portfolio/Weekly/2016/' + ThisWeekDash + ' IG/' + ThisWeekFile
LastWeekPath = 'Y:/Non-US/NB East Asia (Japan & Korea)/Japan Post Bank/IG Credit Portfolio/Weekly/2016/' + LastWeekDash + ' IG/' + LastWeekFile

try:
	os.mkdir(OutPath[:OutPath.rfind('/')])
except WindowsError:
	pass
	#print 'The Path:\n\t{}\nAlready exisits. Please either delete the folder or enter the correct dates'.format(OutPath)

try:
	os.mkdir(OutPath[:OutPath.rfind('/')]+'/Backup')
except WindowsError:
	pass

#other 17.2, uk 4.26, EMU 18.72, US 59.82
############################ EXPOSURE BREAKDOWN ############################
data = p.read_excel(io=DataPath, sheetname='JPIGC-JP IG --Risk and Exp', skiprows=6, skipfooter=3, index_col=0, parse_cols=6, na_values=['', ' '])
TotalDurCon = nm.float64(data.loc[['JPIGCR'], 'Duration Contribution'])
TotalSprdDurCon = nm.float64(data.loc[['JPIGCR'], 'Spread Duration Contribution'])
USDurCon = data[data['Country Name'] == 'United States']
USDurCon.index.name = "US"
EMUDurCon = p.DataFrame(data[data['Country Name'].isin(['Belgium', 'Bulgaria', 'Croatia', 'Cyprus', 'Czech Republic', 'Denmark', 'Estonia', 'Finland', 'France', 'Germany', 'Greece', 'Hungary', 'Ireland', 'Italy', 'Latvia', 'Lithuania', 'Luxembourg', 'Malta', 'Netherlands', 'Poland', 'Portugal', 'Romania', 'Slovakia Republic', 'Slovenia', 'Spain', 'Sweden', 'Switzerland'])])
EMUDurCon.index.name = 'EMU'
UKDurCon = data[data['Country Name'] == 'United Kingdom']
UKDurCon.index.name = "UK"
data.drop(labels=["US", "EMU", "UK", "Other", "JPIGCR"], inplace=True)
OtherDurCon = p.DataFrame(data[~data['Country Name'].isin(['Belgium', 'Bulgaria', 'Croatia', 'Cyprus', 'Czech Republic', 'Denmark', 'Estonia', 'Finland', 'France', 'Germany', 'Greece', 'Hungary', 'Ireland', 'Italy', 'Latvia', 'Lithuania', 'Luxembourg', 'Malta', 'Netherlands', 'Poland', 'Portugal', 'Romania', 'Slovakia Republic', 'Slovenia', 'Spain', 'Sweden', 'Switzerland', 'United States', 'United Kingdom', 'JPHSHY', 'US', 'EMU', 'UK', 'Other'])])
OtherDurCon.index.name = "Other"
DurBuckets = [nm.float64(0.000), nm.float64(0.49999), nm.float64(1.99999), nm.float64(4.9999), nm.float64(6.9999), nm.float64(9.99999), nm.float64(19.9999), nm.float64(100)]
sprdDurBuckets = [nm.float64(0.000), nm.float64(.99999), nm.float64(2.99999), nm.float64(4.9999), nm.float64(6.9999), nm.float64(8.99999), nm.float64(10.9999), nm.float64(100)]
SDHYCountryBuckets = [USDurCon, EMUDurCon, UKDurCon, OtherDurCon]
BucketLabels = ['0-6 month', '6 month-2 year', '2-5 year', '5-7 year', '7-10 year', '10-20year', '20+ year']
sprdBucketLabels = ['0-1 year', '1-3 year', '3-5 year', '5-7 year', '7-9 year', '9-11 year', '11+ year']


def WgtDurCountryTable(group, weight, bin, label, total, country=SDHYCountryBuckets):
	'''group-metric to group by, weight-values to use as weighting, bin-buckets to group by, label-label for the table to include on time axis'''
	tbl = p.DataFrame()
	for i in country:
		wgt = i.groupby(p.cut(i[group], bins=bin, labels=label), group_keys=False).sum()
		wgt = p.DataFrame(wgt[weight] / total)
		wgt = wgt.rename(columns={weight: str(i.index.name)})
		wgt = wgt.fillna(nm.float64(0.00))
		if tbl.empty is True:
			tbl = wgt
		else:
			tbl = tbl.merge(wgt, left_index=True, right_index=True, how='left')
	#tbl.loc[:, 'Total'] = p.Series(nm.sum(tbl, axis=1), index=tbl.index)
	tbl.reset_index(drop=True, inplace=True)
	tbl.index = label
	tbl = tbl.T
	#tbl.loc[:, 'Total'] = p.Series(nm.sum(tbl, axis=1), index=tbl.index)
	#tbl.replace('0','',inplace=True)
	#tbl = tbl.mul(100.0)
	#tbl = per(tbl)
	return tbl

dur = WgtDurCountryTable(group='Duration', weight='Duration Contribution', bin=DurBuckets, label=BucketLabels, total= TotalDurCon)
sprd_dur = WgtDurCountryTable(group='Spread Duration', weight='Spread Duration Contribution', bin=sprdDurBuckets, label=sprdBucketLabels, total = TotalSprdDurCon)

############################ YTW ############################
ReadYTW = p.read_excel(io=DataPath, sheetname='JPIGC-JP IG --Portfolio Gr', skiprows=6, skipfooter=3, index_col=0, parse_cols=1, na_values=['', ' '])

############################ CASH ############################
ReadCashMV = p.read_excel(io=DataPath, sheetname='JPIGC-JP IG --Risk and Exp(1)', skiprows=6, skipfooter=4, index_col=0, parse_cols=2, na_values=['', ' '])

############################ By Sector SDHY ############################
Sectors = ['Industrial', 'Basic Industry', 'Capital Goods', 'Communications', 'Consumer Cyclical', 'Consumer Non-Cyclical', 'Energy', 'Technology', 'Transportation', 'Other Industrial', 'Utility', 'Electric', 'Natural Gas', 'Other Utility', 'Financial Institutions', 'Banking', 'Brokerage Assetmanagers Exchanges', 'Finance Companies', 'Insurance', 'REITS', 'Other Financial', 'Government-Related', 'Agency', 'Local Authority', 'Sovereign', 'Supranational']

ReadLastSec = p.read_excel(io=LastWeekPath, sheetname='II. By sector for IG', skiprows=2, skip_footer=0, index_col=0, header=0, parse_cols='B:G')
ReadLastSec.dropna(axis=0, inplace=True)

ReadSec = p.read_excel(io=DataPath, sheetname='JPIGC-JP IG --Spread Durat', skiprows=6, skipfooter=3, index_col=0, parse_cols='A:F')
ReadSec.fillna(0, inplace=True)

############################ WRITE ############################
#LOAD LAST WEEKS
wb = xl.load_workbook(LastWeekPath)

############################ WRITE EXPOSURE BREAKDOWN ############################
exposurews = wb.get_sheet_by_name('II. Exposure Breakdown')
#DURATION TABLE
durBegRow = 7
durBegCol = 3
for c in range(0, len(dur.columns)):
	for r in dur.iloc[:, c]:
		exposurews.cell(row=durBegRow, column=durBegCol, value=r)
		durBegRow += 1
	durBegCol += 1
	durBegRow = 7
#US
exposurews['C15'] = '=0.00'
exposurews['D15'] = '=0.00'
exposurews['E15'] = '=0.00'
exposurews['F15'] = '=0.00'
exposurews['G15'] = '=0.00'
exposurews['H15'] = '=0.00'
exposurews['I15'] = '=0.00'

exposurews['C16'] = '=IF(ISBLANK(C7),0.00,C7)'
exposurews['D16'] = '=IF(ISBLANK(D7),0.00,D7)'
exposurews['E16'] = '=IF(ISBLANK(E7),0.00,E7)'
exposurews['F16'] = '=IF(ISBLANK(F7),0.00,F7)'
exposurews['G16'] = '=IF(ISBLANK(G7),0.00,G7)'
exposurews['H16'] = '=IF(ISBLANK(H7),0.00,H7)'
exposurews['I16'] = '=IF(ISBLANK(I7),0.00,I7)'

exposurews['C17'] = '=0.00'
exposurews['D17'] = '=0.00'
exposurews['E17'] = '=0.00'
exposurews['F17'] = '=0.00'
exposurews['G17'] = '=0.00'
exposurews['H17'] = '=0.00'
exposurews['I17'] = '=0.00'

exposurews['C18'] = '=0.00'
exposurews['D18'] = '=0.00'
exposurews['E18'] = '=0.00'
exposurews['F18'] = '=0.00'
exposurews['G18'] = '=0.00'
exposurews['H18'] = '=0.00'
exposurews['I18'] = '=0.00'

#EMU
exposurews['C23'] = '=0.00'
exposurews['D23'] = '=0.00'
exposurews['E23'] = '=0.00'
exposurews['F23'] = '=0.00'
exposurews['G23'] = '=0.00'
exposurews['H23'] = '=0.00'
exposurews['I23'] = '=0.00'

exposurews['C24'] = '=IF(ISBLANK(C8),0.00,C8)'
exposurews['D24'] = '=IF(ISBLANK(D8),0.00,D8)'
exposurews['E24'] = '=IF(ISBLANK(E8),0.00,E8)'
exposurews['F24'] = '=IF(ISBLANK(F8),0.00,F8)'
exposurews['G24'] = '=IF(ISBLANK(G8),0.00,G8)'
exposurews['H24'] = '=IF(ISBLANK(H8),0.00,H8)'
exposurews['I24'] = '=IF(ISBLANK(I8),0.00,I8)'

exposurews['C25'] = '=0.00'
exposurews['D25'] = '=0.00'
exposurews['E25'] = '=0.00'
exposurews['F25'] = '=0.00'
exposurews['G25'] = '=0.00'
exposurews['H25'] = '=0.00'
exposurews['I25'] = '=0.00'

exposurews['C26'] = '=0.00'
exposurews['D26'] = '=0.00'
exposurews['E26'] = '=0.00'
exposurews['F26'] = '=0.00'
exposurews['G26'] = '=0.00'
exposurews['H26'] = '=0.00'
exposurews['I26'] = '=0.00'

#UK
exposurews['C31'] = '=0.00'
exposurews['D31'] = '=0.00'
exposurews['E31'] = '=0.00'
exposurews['F31'] = '=0.00'
exposurews['G31'] = '=0.00'
exposurews['H31'] = '=0.00'
exposurews['I31'] = '=0.00'

exposurews['C32'] = '=IF(ISBLANK(C9),0.00,C9)'
exposurews['D32'] = '=IF(ISBLANK(D9),0.00,D9)'
exposurews['E32'] = '=IF(ISBLANK(E9),0.00,E9)'
exposurews['F32'] = '=IF(ISBLANK(F9),0.00,F9)'
exposurews['G32'] = '=IF(ISBLANK(G9),0.00,G9)'
exposurews['H32'] = '=IF(ISBLANK(H9),0.00,H9)'
exposurews['I32'] = '=IF(ISBLANK(I9),0.00,I9)'

exposurews['C33'] = '=0.00'
exposurews['D33'] = '=0.00'
exposurews['E33'] = '=0.00'
exposurews['F33'] = '=0.00'
exposurews['G33'] = '=0.00'
exposurews['H33'] = '=0.00'
exposurews['I33'] = '=0.00'

exposurews['C34'] = '=0.00'
exposurews['D34'] = '=0.00'
exposurews['E34'] = '=0.00'
exposurews['F34'] = '=0.00'
exposurews['G34'] = '=0.00'
exposurews['H34'] = '=0.00'
exposurews['I34'] = '=0.00'

#OTHER
exposurews['C39'] = '=0.00'
exposurews['D39'] = '=0.00'
exposurews['E39'] = '=0.00'
exposurews['F39'] = '=0.00'
exposurews['G39'] = '=0.00'
exposurews['H39'] = '=0.00'
exposurews['I39'] = '=0.00'

exposurews['C40'] = '=IF(ISBLANK(C10),0.00,C10)'
exposurews['D40'] = '=IF(ISBLANK(D10),0.00,D10)'
exposurews['E40'] = '=IF(ISBLANK(E10),0.00,E10)'
exposurews['F40'] = '=IF(ISBLANK(F10),0.00,F10)'
exposurews['G40'] = '=IF(ISBLANK(G10),0.00,G10)'
exposurews['H40'] = '=IF(ISBLANK(H10),0.00,H10)'
exposurews['I40'] = '=IF(ISBLANK(I10),0.00,I10)'

exposurews['C41'] = '=0.00'
exposurews['D41'] = '=0.00'
exposurews['E41'] = '=0.00'
exposurews['F41'] = '=0.00'
exposurews['G41'] = '=0.00'
exposurews['H41'] = '=0.00'
exposurews['I41'] = '=0.00'

exposurews['C42'] = '=0.00'
exposurews['D42'] = '=0.00'
exposurews['E42'] = '=0.00'
exposurews['F42'] = '=0.00'
exposurews['G42'] = '=0.00'
exposurews['H42'] = '=0.00'
exposurews['I42'] = '=0.00'

sprdBegRow = 49
sprdBegCol = 3
for c in range(0, len(sprd_dur.columns)):
	for r in sprd_dur.iloc[:, c]:
		exposurews.cell(row=sprdBegRow, column=sprdBegCol, value=r)
		sprdBegRow += 1
	sprdBegCol += 1
	sprdBegRow = 49

############################ WRITE YTW ############################
YTWws = wb.get_sheet_by_name('II. YTW')
YTWws.cell(row=YTWws.max_row+1, column=1, value=ThisWeek)
YTWws.cell(row=YTWws.max_row, column=2, value=float(ReadYTW.iloc[0]))

############################ WRITE CASH ############################
Cashws = wb.get_sheet_by_name('II. Cash weight')
Cashws['B5'] = ReadCashMV.iloc[0, 0]
Cashws['C5'] = ReadCashMV.iloc[0, 1]

############################ WRITE By Sector IG ############################
#LAST WEEK TABLE VALUES
ws = wb.get_sheet_by_name('II. By sector for IG')
LastSecBegRow = 6
LastSecBegCol = 8
for c in range(0, len(ReadLastSec.columns)):
	for r in ReadLastSec.iloc[:-1, c]:
		ws.cell(row=LastSecBegRow, column=LastSecBegCol, value=r)
		LastSecBegRow += 1
	LastSecBegCol += 1
	LastSecBegRow = 6

#LAST WEEK TOTAL
LastTotalRow = 33
LastTotalCol = 8
for t in ReadLastSec.iloc[-1]:
	ws.cell(row=LastTotalRow, column=LastTotalCol, value=t)
	LastTotalCol += 1

#THIS WEEK TABLE VALUES
SecBegRow = 6
SecBegCol = 3
for c in range(0, len(ReadSec.columns)):
	for r in ReadSec.iloc[1:, c]:
		ws.cell(row=SecBegRow, column=SecBegCol, value=r)
		SecBegRow += 1
	SecBegCol += 1
	SecBegRow = 6

#THIS WEEK TOTAL
TotalRow = 33
TotalCol = 3
for t in ReadSec.iloc[0]:
	ws.cell(row=TotalRow, column=TotalCol, value=t)
	TotalCol += 1

#WRITE DATES AT TOP
ws.cell(row=2, column=8, value=LastWeek)
ws.cell(row=2, column=3, value=ThisWeek)

wb.save(OutPath)
Backupwb = xl.load_workbook('Y:/Non-US/NB East Asia (Japan & Korea)/Japan Post Bank/IG Credit Portfolio/Weekly/2016/Backup/II. Exposure Breakdown IG Backup '+ThisWeekDash+'.xlsx')
Backupwb.save(OutPath[:OutPath.rfind('/')] + '/Backup/' + 'II. Exposure Breakdown IG Backup '+ThisWeekDash+'.xlsx')
