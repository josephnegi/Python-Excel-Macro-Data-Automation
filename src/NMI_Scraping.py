# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup

comments=[]
paragraphs = {}
pg=[]
with open(r'C:/Users/joeyn/Documents/Indicators/US/ISM Reports/NMI/September NMI.html', 'r', encoding='utf-8') as html_file:
    content = html_file.read()
    soup = BeautifulSoup(content,'lxml')
    tags = soup.find_all('li')
    for i, v in enumerate(tags):        
        #print(type(tags[i].text))
        if '[' in tags[i].text:
            #print(True)
            #print(i, ":", tags[i].text)
            comments.append(tags[i].text)
        #print()
    t=[]
    #print(comments)
    para1 = soup.find_all('div', class_='col-lg-12')
    for para in para1:
        ps = para.find_all('p')
        #print(ps[-2].text)
        #print()
        t.append(ps[-2].text)
    #print(t.pop(0))
    pg.append(t.pop(0))
    paragraphs['ISM Non-Manufacturing'] = str(pg[0]) 
    #pg.append(str(paragraphs['ISM Non-Manufacturing']))
    #print(paragraphs)
    paras = soup.find_all('div', class_='row')
    for i, x in enumerate(paras):
        if i <18:
            continue
        elif i<28:
           # print('Index',i)
            #print(x.h3.text)
            #print()
            ps = x.find_all('p')
            #print(ps[-1].text)
            paragraphs[x.h3.text] = ps[-1].text
            pg.append(ps[-1].text)
            #print('================================================')
    
    #printing the paras
    for i, v in paragraphs.items():
        print(i, ":", v)
        print()
    #[print(pg[x], end='\n\n') for x,v in enumerate(pg)]

