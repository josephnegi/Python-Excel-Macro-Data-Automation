# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup

comments=[]
paragraphs = {}
pg=[]
with open('C:/Users/joeyn/Documents/Indicators/US/ISM Reports/PMI/September PMI.html', 'r') as html_file:
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
    para1 = soup.find_all('div', class_='col-lg-12')
    for para in para1:
        ps = para.find_all('p')
        #print(ps[-1].text)
        #print()
        t.append(ps[-1].text)
    #print(comments)

    paragraphs['ISM Manufacturing'] = str(t.pop(0)) 
    pg.append(str(paragraphs['ISM Manufacturing']))
    #print(paragraphs)
    paras = soup.find_all('div', class_='mb-4')
    for i, x in enumerate(paras):
        if i == 0:
            continue
        elif i <9:
            #print('Index',i)
            #print(x.h3.text.strip('*'))
            #print()
            ps = x.find_all('p')
            #print(ps[-1].text)
            paragraphs[x.h3.text.strip('*')] = ps[-1].text
            pg.append(ps[-1].text)
            #print('================================================')
    
    #printing the paras
    for i, v in paragraphs.items():
        print(i, ":", v)
        print()
    #[print(pg[x], end='\n\n') for x,v in enumerate(pg)]