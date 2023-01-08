# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
#%%---------------- Import Libraries--------------------------

import pandas as pd
import numpy as np
#import pyodbc
import datetime as dt
from pathlib import Path

import re
import win32com.client as win32
import os





#%%---------------- Setting up Date Parameters--------------------------

date = dt.date.today()

week= date.strftime("%Y%W")

start_week = date 
end_week = start_week - dt.timedelta(8)

#%%---------------- Import DCG file and file cleaning--------------------------

Hi_Low=pd.read_excel("//itseelm-nt0042/Common_L/IOS-SC Planning Support/NEWS/NEWS_DCG SET UP/TEMPLATE_NEWS_DCG SET UP TO NP.xlsx", "Proposal")

Reference= pd.read_excel("//ITSEELM-NT0044/MUJAM1$/Desktop/IKEA folder/Lead responsibilities/Hi Low Flow/Codes/Employees.xlsx", "Sheet1")

Hi_Low=Hi_Low.drop([0,1,3], axis=0, inplace= False)

Hi_Low.columns = Hi_Low.iloc[0]

Hi_Low=Hi_Low.drop([2], axis=0, inplace= False)



Hi_Low['ARTS ADDED'] = pd.to_datetime(Hi_Low['ARTS ADDED']).dt.date

#%%---------------- Filter on date added in last 8 days--------------------------

Hi_Low=Hi_Low[Hi_Low['ARTS ADDED'] >= end_week]

#%%----------Merge Contact---------

Merged = pd.merge(Hi_Low, Reference, on ='Need Planner')

#%%----------Creating a new folder for every new week---------
if not os.path.exists('//ITSEELM-NT0044/MUJAM1$/Desktop/IKEA folder/Lead responsibilities/Hi Low Flow/FRD files/'+week): 
    os.makedirs('//ITSEELM-NT0044/MUJAM1$/Desktop/IKEA folder/Lead responsibilities/Hi Low Flow/FRD files/'+week)


#%% ---------- saving file per NP and Sending out emails-------------------
    
for emailid in Merged['Email'].unique():
    temp_Hi_Low=Merged.loc[(Merged['Email']== emailid)]
    if len(temp_Hi_Low) != 0:
        temp_Hi_Low.to_excel("//ITSEELM-NT0044/MUJAM1$/Desktop/IKEA folder/Lead responsibilities/Hi Low Flow/FRD files/"+ str(week)+"/"+ str(emailid)+'.xlsx', index=False)
    
    
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = emailid
        #mail.CC= 'javier.fernandez@renault-trucks.com'
        #mail.BCC= 'georges.da.fonte@volvo.com'
        mail.Subject = 'Hi_Low Flow Additions' 
        mail.Body = 'Message body'
    
       
        mail.HTMLBody = '''<Font Size = 4 Face= Calibri Color=#2A1E5F>
    
        Dear Need Planner,<br><br><br>
        Please find enclosed report for Hi_Low flow change to the articles that have been included in last week <br><br><br>
    
        If you want to give your comments for any change kindly go to the file through below link
        
        
        <br><br>
         
        <a href="\\\\itseelm-nt0042\\Common_L\\IOS-SC Planning Support\\NEWS\\NEWS_DCG SET UP">News DCG SETUP FILE</a>
        <br><br> If this doesnt work then click on below: <br><br>
        "\\\\itseelm-nt0042\\Common_L\\IOS-SC Planning Support\\NEWS\\NEWS_DCG SET UP"
        
      
         
        
         
        
        <br><br>Kind regards <br>
            
        Muhammad Aarish Jamil<br>
        Need Planner (BA Tesec)<br>
            

        <br><br></font>
        '''  
        
        mail.Attachments.Add("//ITSEELM-NT0044/MUJAM1$/Desktop/IKEA folder/Lead responsibilities/Hi Low Flow/FRD files/"+ str(week)+"/"+ str(emailid)+'.xlsx')
        
        #mail.Send()
        mail.Display()

#