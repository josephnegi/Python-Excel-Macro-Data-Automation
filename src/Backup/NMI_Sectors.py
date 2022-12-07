# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd
from datetime import datetime, timedelta, date
import xlwings as xw
#Functions
def op(s):
    ss=s
    if ':' in s:
        st = s.split(': ')
        ss = st[0]
        s1 = st[1]
    elif '(' in s and ')' in s:
        s1 = s[s.rfind('(')+1:s.rfind(')')]
    print(ss)#Prints out the heading less the industries
    strlist = [ss]
    res = [sub.split() for sub in strlist]
    words = res[0]#words is a list containing all words of the sentence
    t_list=[]
    _gc=0
    #Fiding the number mentioned in the sentence
    for ind, val in enumerate(words):
        for i, v in nos.items():
            vc = "% s" % v
            if i == val.lower() or vc == val.lower():
                t_list.append(v)
    #Finding if there's growth or contraction _gc=1 is grwth _gc=0 is contrac^n           
        for i, v in gc.items():
            if i == val.lower():
                _gc += v
                
    t_list.sort()
    print(t_list)
    #print(_gc)
    no = t_list[0]
    #i_list = s1[1].split(';') 
    return listOut(no, s1, _gc)
    #ask is this correct?([y]/n)
        #if yes, sort the df then send for concatenation
    
def listOut(num, tt, gc_):
    i_list=[]
    if tt != '':
        i_list = tt.split(';')
        #this for cleans the list for any unwanted characters
        for ind, val in enumerate(i_list):
            for i,v in Industries.items():
                if i in val:
                    i_list[ind] = i
    if num != len(i_list) or tt=='':
        print('There is a change in the input of the Header text and industries')
        y= input('Do you want to enter the details manually? (y/n)')
        if y == 'y':
            n1 = input('Enter the no of industries:')
            i1 = input('Enter the industries:')
            i1_list = i1.split(';')
            #this for cleans the list for any unwanted characters
            for ind, val in enumerate(i1_list):
                for i,v in Industries.items():
                    if i in val:
                        i1_list[ind] = i
            gc1 = int(input('Enter 1 for growth or 0 for contraction?:'))
            return makeDataframe(n1, i1_list, gc1)
    else:
        return makeDataframe(num, i_list, gc_)

def makeDataframe(n, _list, _gc):
    
    lt = pd.Series(_list) 
    if _gc == 0:
        arr = np.arange(-len(lt),0,1)  
    elif _gc == 1:
        arr = np.arange(len(lt),0,-1)
    df1 = pd.DataFrame([lt, arr])
    df2 = df1.transpose()
    df2.columns =['Industries', 'Rank'] 

    print(df2)
    ask = input('Is the above data correct? (y/n)')
    if ask == 'y':
        return df2
    elif ask =='n':
        return listOut(None, '', None)
    
def getData(s):
    data=pd.DataFrame()
    for i, v in enumerate(s):
        #print(i)
        if i==0:
            data1 = op(v)
        elif i==1:
            data2 = op(v)
    data = pd.concat([data1, data2], sort=True)
    ar = np.arange(0,len(data))
    data.index = ar
    print(data)
    return data    
ind = ['ISM Non-Manufacturing', 'Business Activity','New Orders', 'EMPLOYMENT', 'DELIVERIES',
       'INVENTORIES']
nos={
     'one':1, 'two':2, 'three':3, 'four':4, 'five':5, 'six':6, 'seven':7,
     'eight':8, 'nine':9, 'ten':10, 'eleven':11, 'twelve':12, 'thirteen':13, 
     'fourteen':14, 'fifteen':15, 'sixteen':16, 'seventeen':17, 'eighteen':18
     }
gc = {'growth':1, 'decline':0, 'increase':1, 'decrease':0, 'higher':1,
     'lower':0, 'high':1, 'low':0, 'contraction':0, 'slower':0, 'faster':1
     }
Industries={
    'Agriculture, Forestry, Fishing & Hunting': 1, 
    'Mining': 2, 
    'Utilities': 3, 
    'Construction': 4, 
    'Wholesale Trade': 5, 
    'Retail Trade': 6, 
    'Transportation & Warehousing': 7, 
    'Information': 8, 
    'Finance & Insurance': 9, 
    'Real Estate, Rental & Leasing': 10, 
    'Professional, Scientific & Technical Services': 11, 
    'Management of Companies & Support Services': 12, 
    'Educational Services': 13, 
    'Health Care & Social Assistance': 14, 
    'Arts, Entertainment & Recreation': 15, 
    'Accommodation & Food Services': 16, 
    'Public Administration': 17, 
    'Other Services': 18
    }

wb = xw.Book(r'../Indicators/US/ISM Reports/ISM_NonManufacturing_Index.xlsx')
sht = wb.sheets[13]
lcol = sht.range((7,3)).end('right').column
print(lcol)
#Step 2a: Getting the new dates to input in the sheet
last_month = datetime.today().replace(hour=0, second=0, microsecond=0, minute=0) - timedelta(days=datetime.today().day)
last_month = last_month - timedelta(days=last_month.day-1)

#Step 1b: Copying the last month's values to the next month
sht.range((1,lcol-1), (161, lcol)).copy(sht.range((1,lcol+1),(1,lcol+2)))

#Step2b: loop through all descriptions and change the value to 0
for i, v in enumerate(ind):
    row=None
    #finding the row no of the heading
    for x in range(1, 161):
        t = str(sht.range((x,2)).value)
        if t.upper() == v.upper():
            print(sht.range((x,2)).value)     
            row = x
    print(row)
    #Checking if the date values are already updated
    if sht.range((row,lcol-1)).value == last_month:
        sht.range((row, lcol+1), (row+21, lcol+2)).delete(shift='left')
        print(f'The data for {v} has already been updated for this month!')
        continue
    
    else:
        #changing the date of the new heading
        sht.range((row,lcol+1)).value = last_month
        #changing the default values to 0
        for x in range(row+3, row+21):
            sht.range((x,lcol+2)).value = 0
        #taking the input for the details para|loop through this later using scraping
        para = input(f'Enter the data for {v}:')
        a = para.split('. ',maxsplit=1)
        df = getData(a)
        
        #filling in the actual values
        for x in range(row+3, row+21):
            for y in range(0,len(df)):
                #print(df4.iloc[y,0])
                if df.iloc[y,0] == sht.range((x,2)).value:    
                    sht.range((x,lcol+2)).value = df.iloc[y,1]
    
#input('Enter the String:')
#para='Nine manufacturing 18 industries reported growth in September, in the following order: Nonmetallic Mineral Products; Machinery; Plastics & Rubber Products; Miscellaneous Manufacturing; Apparel, Leather & Allied Products; Transportation Equipment; Food, Beverage & Tobacco Products; Computer & Electronic Products; and Electrical Equipment, Appliances & Components. The seven industries reporting contraction in September compared to August, in the following order are: Furniture & Related Products; Textile Mills; Wood Products; Printing & Related Support Activities; Paper Products; Chemical Products; and Fabricated Metal Products.'
#print(para)


#para.find('.')
#o = para.partition('.')
#print(o[0])
#a=para.split('. ',maxsplit=1)
#print(a)

#df=getData(a)

#def analyse(f_half, s_half):
    #f1 = f_half.partition(':')
    #header_text = GrowthContraction(f1[0])
 


