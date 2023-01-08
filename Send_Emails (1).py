# -*- coding: utf-8 -*-
"""
Created on Mon Mar 16 09:42:01 2020

@author: a309695
"""

import pandas as pd
import numpy as np
import pyodbc
import datetime as dt

import re
import win32com.client as win32
import os
#from datetime import date, timedelta 
#ExcelApp=win32.GetActiveObject("Excel.Application")

#ExcelWrkbook=ExcelApp.Workbooks('TEMPLATE_NEWS_DCG_SET_UP_TO_NP.xlsx')


#aaa=ExcelWrkbook.path
#aaa=aaa.replace('\\','/')

#'Book1.xlsm'
#ab=pd.read_excel(aaa+'/'+ ExcelWrkbook.name )

#------------ pre-requisite Connect to DSP info---------------------------

#a=dt.datetime.today()-dt.timedelta(days=14)
#a=today.strftime('%m').lstrip('0')

#Dealer_count=pd.read_excel("//vcn.ds.volvo.net/parts-got/proj02/003742/DIM/Analyses/push variable supersession/Digit_information.xlsx")



today = dt.datetime.today()





week_prior =  today - timedelta(days=7)

Hi_Low=pd.read_excel("//ITSEELM-NT0044/MUJAM1$/Desktop/IKEA folder/Lead responsibilities/Hi Low Flow/Codes/TEMPLATE_NEWS_DCG_SET_UP_TO_NP.xlsx", "Proposal")

Hi_Low= Hi_Low.iloc[2:]

new_header = Hi_Low.iloc[0] #grab the first row for the header
Hi_Low = Hi_Low[1:] #take the data less the header row
Hi_Low.columns = new_header #set the header row as the df header

Hi_Low= Hi_Low.iloc[2:]

Hi_Low['creation week']=  pd.to_datetime(Hi_Low['ARTS ADDED'])

Sliced=Hi_Low.loc[Hi_Low['ARTS ADDED'] > week_prior ]

"""

Red_Card_List=pd.read_excel(aaa+'/'+ ExcelWrkbook.name )

Red_Card_List['COUNTRY']=Red_Card_List['COUNTRY'].fillna('Not Applicable')

Red_Card_List=Red_Card_List.loc[ (Red_Card_List['Pushed Qty.']> 0) & (Red_Card_List['COUNTRY'] != 'Not Applicable')]


Red_Card_List['Company']=Red_Card_List['Company'].astype(str)

Red_Card_List['Company']=Red_Card_List['Company'].apply(lambda x: x.zfill(2))
To_be_sent=Red_Card_List[['Company','District','Dealer','Prefix', 'Part', 'Pushed Qty.', 'COUNTRY' ]].astype(str)
#To_be_sent=To_be_sent.rename(columns={'part to push':'Part', 'qty pushed':'Pushed Qty.', 'COMPANY CODE': 'Company','DISTRICT': 'District' ,'DEALER':'Dealer', 'PREFIX':'Prefix' }, inplace=False)

Dealer_zero=pd.merge(To_be_sent, Dealer_count, on='COUNTRY')

if not os.path.exists(aaa+'/Push_Part_Files'): 
    os.makedirs(aaa+'/Push_Part_Files')
    
for countrycode in To_be_sent['COUNTRY'].unique():
    List=To_be_sent.loc[To_be_sent['COUNTRY']== countrycode]
    abc=countrycode
    machine=Dealer_count.loc[Dealer_count['COUNTRY']== countrycode ].reset_index(drop=True)
       
    if not machine.empty:
        machine=machine.iloc[0,1]
        List=List.astype(str)
        
        List['Dealer']=List['Dealer'].astype(str).str.zfill(machine) 
    del List['COUNTRY']
    List.to_csv(aaa+'/Push_Part_Files/'+abc+'_PUSHED.csv', index=False, sep=';')
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'dsp@b2b.volvo.com'
    
    mail.Subject = 'R3NI4T'+abc 
       
    mail.Attachments.Add(aaa+'/Push_Part_Files/'+abc+'_PUSHED.csv')
    
    #mail.Send()
    mail.Display()
"""