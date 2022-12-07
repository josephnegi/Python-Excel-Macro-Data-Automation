# -*- coding: utf-8 -*-
"""
Created on Fri Oct 28 17:07:36 2022
Program to Automate the PMI and NMI Spreadsheet for GDP and ^GSPC

@author: joeyn
"""
import requests_cache
import pandas_datareader as web
from pandas_datareader.yahoo.headers import DEFAULT_HEADERS
import numpy as np
import pandas as pd
from datetime import datetime, timedelta, date
import xlwings as xw
import time

#Functions
def DatesOfReleases(d):
    yr = datetime.today().year
    if d.month == 1:
        return datetime(yr, 3, 31, hour=0, minute=0, second=0, microsecond=0)
    elif d.month == 4:
        return datetime(yr, 6, 30, hour=0, minute=0, second=0, microsecond=0)
    elif d.month == 7:
        return datetime(yr, 9, 30, hour=0, minute=0, second=0, microsecond=0)
    elif d.month == 10:
        return datetime(yr, 12, 31, hour=0, minute=0, second=0, microsecond=0)
def Assigning(a):
    try:
        t = float(gspcdf.loc[a,'Adj Close'])
        return t
        return 1
    except KeyError:
        return 0
    except:
        return -1

#Caching
expire_after = timedelta(days=1)
session1 = requests_cache.CachedSession(cache_name='cache1', backend='sqlite', expire_after=expire_after)
session2 = requests_cache.CachedSession(cache_name='cache2', backend='sqlite', expire_after=expire_after)
session2.headers = DEFAULT_HEADERS

'''Automating the PMI Sheet First'''

            #getting current real GDP values for last 5 years
gdp = web.DataReader('GDPC1', data_source='fred', pause=0.2, session=session1)
#Storing the Value and Date in Variables
gdp_val = gdp.iloc[-1]
gdp_dt = gdp.iloc[-1].name #- timedelta(days=1)
gdp_date = DatesOfReleases(gdp_dt)
#This gets the GDP value as per the last
#date of the month, like in the spreadsheets

            #Storing the variables in the spreadsheet
#opening the sheet
wb = xw.Book(r'../Indicators/US/ISM Reports/ISM_Manufacturing_SP500_GDP.xlsx')
length = len(wb.sheets) #Getting the no of sheets in the workbook
sht1 = wb.sheets[0]

#Writing the GDP Value First. Will be checked every month
#df1 = sht1[0,0].options(pd.DataFrame, expand='table').value
len_sht = sht1[0,0].end('down').row
#Finding the cell value of the Index (y)
for i in range(len_sht-10, len_sht+1):
    if sht1.range((i,1)).value == gdp_date:
        row = i
        break
    elif sht1.range((i+1, 1)).value == None:
        row = i+1
        sht1.range((row, 1)).value = gdp_date
#writing the value of the GDP into the cell
sht1.range((row, 3)).value = gdp_val.values
print("GDP Values Updated")

            #Writing the values of the PMI
#getting the data
flag = False
while flag == False:
    try:
        pmi_html = str(input('Provide the path of the recent month\'s pmi HTML Report:'))
        pmidf = pd.read_html(pmi_html)
        flag = True
    except:
        print('There is an error in the inputs, try again')
pmi = pmidf[0]#Pandas Series which stores the first table of all pmi month data
pmi_val = float(pmi.iloc[0,1])
#Getting the date of last month for the pmi value
last_month = datetime.today().replace(hour=0, second=0, microsecond=0, minute=0) - timedelta(days=datetime.today().day)
#writing the value of the PMI into the cell
for i in range(len_sht-10, len_sht+1):
    if sht1.range((i,1)).value == last_month:
        row = i
        break
    elif sht1.range((i+1, 1)).value == None:
        row = i+1
        sht1.range((row, 1)).value = last_month
        
#Writing the value in the cell
sht1.range((row,2)).value = pmi_val
print("PMI Values Updated")

                #Getting the ^GSPC closing date and value of last month
gspcdf = web.DataReader('^GSPC', data_source='yahoo', pause=0.2, session=session2)
cdt = last_month #closing date to be used to fetch the GSPC Values
gspc_val = 0.0
flag=False #Flag to run the exception
while flag == False:
    T = Assigning(cdt)
    if T >= 1:
        gspc_val = T
        flag = True
    elif T == 0:
        cdt = cdt - timedelta(days=1)
    elif T == -1:
        print("There has been an UNKNOWN error, check the APIs and/or Code!")

#writing the gspc Value in the cells
#Switching the sheets
sht2 = wb.sheets[1]
#get the row cell
len_sht2 = sht2[0,0].end('down').row
for i in range(len_sht2-6, len_sht2+1):
    #print(sht2.range((i,1)).value == last_month)
    if sht2.range((i,1)).value == last_month:
        row = i
        break
    elif sht2.range((i+1, 1)).value == None:
        row = i+1
        sht2.range((row, 1)).value = last_month
#writing the values
sht2.range((row,2)).value = gspc_val
sht2.range((row,5)).value = pmi_val
print("^GSPC Values Updated")

wb.save()
print("Closing in 3 Seconds")
time.sleep(3)
wb.close()
print(r'Workbook1 ISM_Manufacturing_SP500_GDP.xlsx Saved and Closed')


'''Automating the NMI Sheet Second'''
print(r'Starting Sheet 2 ISM_NonManufacturing_SP500_GDP.xlsx')
wb2 = xw.Book(r'../Indicators/US/ISM Reports/ISM_NonManufacturing_SP500_GDP.xlsx')
sht1 = wb2.sheets[0]

#Getting the Values of The NMI
flag = False
while flag == False:
    try:
        nmi_html = str(input('Provide the path of the recent month\'s nmi HTML Report:'))
        nmidf = pd.read_html(nmi_html)
        flag = True
    except:
        print('There is a error in the input, try again')
nmi = nmidf[0]
nmi_val = float(nmi.iloc[0, 1])


#finding pmi date cell
len_sht1 = sht1[0,0].end('down').row
for i in range(len_sht1-10, len_sht1+1):
    if sht1.range((i,1)).value == last_month:
        row = i
        break
    elif sht1.range((i+1, 1)).value == None:
        row = i+1
        sht1.range((row,1)).value = last_month
        #Assign the row date as well
#replacing the values
sht1.range((row,2)).value = nmi_val
for i in range(len_sht1-10, len_sht1+1):
    if sht1.range((i,1)).value == gdp_date:
        row = i
        break
    elif sht1.range((i+1, 1)).value == None:
        row = i+1
        sht1.range((row, 1)).value = gdp_date
#writing the value of the GDP into the cell
sht1.range((row, 3)).value = gdp_val.values
print('Sheet 1 Values updated')

#writing sheet 2 data
sht2 = wb2.sheets[1]
len_sht2 = sht2[0,0].end('down').row
for i in range(len_sht2-6, len_sht2+1):
    if sht2.range((i,1)).value == last_month:
        row = i
        break
    elif sht2.range((i+1, 1)).value == None:
        row = i+1
        sht2.range((row,1)).value = last_month
sht2.range((row,2)).value = gspc_val
sht2.range((row,5)).value = nmi_val
print('Sheet 2 Values Updated')

#writing sheet 3 data
sht3 = wb2.sheets[2]
len_sht3 = sht3[0,0].end('down').row
ba_val = float(nmi.iloc[1, 1])
for i in range(len_sht3-6, len_sht3+1):
    if sht3.range((i,1)).value == last_month:
        row = i
        break
    elif sht3.range((i+1, 1)).value == None:
        row = i+1
        sht3.range((row,1)).value = last_month
        #Assign the row date as well
sht3.range((row,2)).value = ba_val
for i in range(len_sht3-6, len_sht3+1):
    if sht3.range((i,1)).value == gdp_date:
        row = i
        break
    elif sht3.range((i+1, 1)).value == None:
        row = i+1
        sht3.range((row, 1)).value = gdp_date
#writing the value of the GDP into the cell
sht3.range((row, 3)).value = gdp_val.values
print('Sheet 3 Values Updated')

#Writing sheet 4 data
sht4 = wb2.sheets[3]
len_sht4 = sht4[0,0].end('down').row
for i in range(len_sht4-6, len_sht4+1):
    if sht4.range((i,1)).value == last_month:
        row = i
        break
    elif sht4.range((i+1, 1)).value == None:
        row = i+1
        sht4.range((row,1)).value = last_month
sht4.range((row,2)).value = gspc_val
sht4.range((row,5)).value = ba_val
print('Sheet 4 Values Updated')

wb2.save()
print("Closing in 3 Seconds")
time.sleep(3)
wb2.close()
print(r'Workbook2 ISM_NonManufacturing_SP500_GDP.xlsx Saved and Closed')
