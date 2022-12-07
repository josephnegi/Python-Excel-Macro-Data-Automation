# -*- coding: utf-8 -*-
#import requests_cache
#import pandas_datareader as web
#from pandas_datareader.yahoo.headers import DEFAULT_HEADERS
import numpy as np
import pandas as pd
from datetime import datetime, timedelta, date
import xlwings as xw
import time
#import calendar
#functions 
'''
def add_months(sourcedate, months):
    month = sourcedate.month - 1 + months
    year = sourcedate.year + month // 12
    month = month % 12 + 1
    day = min(sourcedate.day, calendar.monthrange(year,month)[1])
    return datetime.date(year, month, day)
'''
#sht->sheet reference; val->Val to be written; l_month-> date to write;w_col->column to be written
#valColDict -> Dict to provide the value(s):column(s) to be printed for the date row
def WriteVal(sht, l_month, valColDict):
    #function that writes the value in the sheets
    len_sht = sht[0,0].end('down').row
    for i in range(len_sht-6, len_sht+1):
        row=None
        if sht.range((i, 1)).value == l_month:
            row = i
            break
        elif sht.range((i+1, 1)).value == None:
            row = i+1
            sht.range((row, 1)).value = l_month
    #sht.range((row, w_col)).value = val
    for v, c in valColDict.items():    
        sht.range((row, c)).value = v
    print('Sheet Updated:', sht.name)
    
wb = xw.Book(r'../Indicators/US/ISM Reports/ISM_Manufacturing_Index.xlsx')
pmidf = pd.read_html('C:/Users/joeyn/Documents/Indicators/US/ISM Reports/PMI/September PMI.html')
pmi=pmidf[0]
print('Today\'s date:', date.today())
last_month = datetime.today().replace(hour=0, second=0, microsecond=0, minute=0) - timedelta(days=datetime.today().day)
last_month = last_month - timedelta(days=last_month.day-1)
print('Input Date as per the spreadsheet:', last_month)
print(last_month.month, last_month.year)
#The second line makes the month end date into month start date

#Sheet Name PMI
sht1= wb.sheets[2]
pmi_val = float(pmi.iloc[0,1])
print(sht1.name, pmi_val)
val_col={pmi_val:2}
WriteVal(sht1, last_month, val_col)
val_col.clear()#Clears the dictionary for the next loop to use the dictionary

#loop to write the sheet names from BA to Inventory
for x in range(1, 11):
    shtx=wb.sheets[x+2]
    x_val = float(pmi.iloc[x,1]) 
    print(shtx.name, x_val)
    #we are using two separate scenarios for exports and imports, as the import
    #sheet is different in comparison to the rest, and needs export values
    if shtx.name == 'Exports':
        exports_val = x_val
        val_col={x_val:2, pmi_val:4}
        WriteVal(shtx, last_month, val_col)
        val_col.clear()
    elif shtx.name == 'Imports':
        val_col={x_val:2, exports_val:4, pmi_val:6}
        WriteVal(shtx, last_month, val_col)
        val_col.clear()
    else:
        val_col={x_val:2, pmi_val:4}
        WriteVal(shtx, last_month, val_col)
        val_col.clear()

wb.save()
print('')
print("Closing in 3 Seconds")
time.sleep(3)
wb.close()
print(r'ISM_Manufacturing_Index.xlsx Saved and Closed')

