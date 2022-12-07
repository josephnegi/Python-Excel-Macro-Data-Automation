# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd
from datetime import datetime, timedelta, date
import xlwings as xw
from bs4 import BeautifulSoup
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
            print('From the previous incorrect segment^')
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
ind = ['ISM Non-Manufacturing', 'Business Activity','New Orders', 'Employment', 'Supplier Deliveries',
       'Inventories']
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
########################################################################
comments=[]
paragraphs = {}
pg=[]
#getting the data
flag = False
while flag == False:
    try:
        nmi_html = str(input('Provide the path of the recent month\'s pmi HTML Report:'))
        #pmidf = pd.read_html(r'C:/Users/joeyn/Documents/Indicators/US/ISM Reports/NMI/September NMI.html', 'r', encoding='utf-8')
        
        with open(nmi_html, 'r', encoding='utf-8') as html_file:
            content = html_file.read()
            soup = BeautifulSoup(content,'lxml')
            tags = soup.find_all('li')
            for i, v in enumerate(tags):        
        
                if '[' in tags[i].text:
        
                    comments.append(tags[i].text)
            t=[]
            #print(comments)
            para1 = soup.find_all('div', class_='col-lg-12')
            for para in para1:
                ps = para.find_all('p')
        
                t.append(ps[-2].text)
        
            pg.append(t.pop(0))
            paragraphs['ISM Non-Manufacturing'] = str(pg[0]) 
        
            paras = soup.find_all('div', class_='row')
            for i, x in enumerate(paras):
                if i <18:
                    continue
                elif i<28:
                    ps = x.find_all('p')
                    paragraphs[x.h3.text] = ps[-1].text
                    pg.append(ps[-1].text)
                    
        flag = True
    except:
        print('There is an error in the inputs, try again')
          

########################################################################
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
        #para = input(f'Enter the data for {v}:')
        para = paragraphs[v]
        print(para)
        print()
        a = para.split('. ',maxsplit=1)
        df = getData(a)
        
        #filling in the actual values
        for x in range(row+3, row+21):
            for y in range(0,len(df)):
                #print(df4.iloc[y,0])
                if df.iloc[y,0] == sht.range((x,2)).value:    
                    sht.range((x,lcol+2)).value = df.iloc[y,1]
    
##########################################################################
#Page to write the Industry comments 
sht1 = wb.sheets[14]
lcol1 = sht1.range((1,1)).end('right').column
print("lcolumn",lcol1)
#sht1.range((1,lcol1)).value == last_month
sht1.range((1,lcol1+1)).value = last_month
for i, v in enumerate(comments):
  
    _ind = v[v.rfind('[')+1:v.rfind(']')]
    _com = v[1:v.rfind('[')-2]

    row = None
    #finding the row no of the heading
    for x in range(2, 20):
        t = str(sht1.range((x,1)).value)
        if t.upper() == _ind.upper():
            print(sht1.range((x,1)).value)     
            row = x
    print('Row:',row)
    #Checking if the date values are already updated
    if sht1.range((row, lcol1+1)).value != None:
        sht1.range((1,lcol1+1)).delete(shift='left')
        print(f'The data {v} has already been updated for this month!')
        continue
    else:
        sht1.range((row,lcol1+1)).value = _com
        print(f'Comment {v} Updated')
  

wb.save()
