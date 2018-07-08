# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import pandas as pd
from requests_html import HTMLSession
import time
from openpyxl import load_workbook

session = HTMLSession()
lst2=[]

url2='http://www.wsj.com/mdc/public/page/2_3062-shtnyse_0_9-listing.html'
r2 = session.get(url2)    
print(r2)
hd=r2.html.find('.colhead')


for x in range(7):
    lst2.append(hd[x].text)


for x in range(7,9):
    out = []
    buff = []  

    for c in hd[x].text:
        if c == '\n':
            out.append(''.join(buff))
            buff = []
        else:
            buff.append(c)
    else:
        if buff:
           out.append(''.join(buff))
           
    lst2.append(' '.join(out[0:len(out)]))
    
time.sleep(1)

headers = lst2
exchange=['shtnyse_', 'shtnasdaq_','shtamex_']
sheet=['NYSE', 'NASDAQ', 'AMEX']
data=['0_9','A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']

for j in range(len(exchange)):
    lst1=[]
    for i in range(len(data)):
        url='http://www.wsj.com/mdc/public/page/2_3062-'+ exchange[j] + data[i] + '-listing.html'
        r = session.get(url)        
        print(r)
        
        symbol=r.html.find('tr')

        for x in range(3,len(symbol)-5):
            out = []
            buff = []
            
            for c in symbol[x].text:
                if c == '\n':
                    out.append(''.join(buff))
                    buff = []
                else:
                    buff.append(c)
            else:
                if buff:
                   out.append(''.join(buff))
            lst1.append(tuple(out[0:len(out)]))
        time.sleep(1)
        
    df=pd.DataFrame(lst1, columns=headers)

    if j < 1:
        writer = pd.ExcelWriter('short.xlsx')
        df.to_excel(writer, sheet[j])  
        writer.save()
        
    else:
        writer = pd.ExcelWriter('short.xlsx', engine='openpyxl')
        book = load_workbook('short.xlsx')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        df.to_excel(writer, sheet[j])  
        writer.save()


