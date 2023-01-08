# -*- coding: utf-8 -*-
"""
Created on Mon Sep 23 14:59:47 2019

1: Get all orders from SIPRE
2: Get all parts with on order in DSP
3: Take orders from SIPRE that are within the lead time of DSP
4: Aggregate parts on order of SIPRE
5: See If Part quantity SIPRE <  to parts on order in DSP then it is OLD on order


@author: a309695
"""
#%%---------Data from QVD files of SIPRE and partslinq
#------------------------------------------
import os
import pandas as pd
import datetime as dt
import numpy as np
import pyodbc
import math 
import win32com.client as win32
sipre=pd.read_csv('//vcn.ds.volvo.net/parts-got/proj02/003742/DIM_Analytics_EMEA_and_APAC/07-Process_Automatization/18_RT_Old_on_Order/Partslinq/OOO.csv')

                  


#%%---------------- Import DSP data from Library R3F002--------------------------

#library1 = 'G1R3APD'
#library2 = 'G1R3BPD'
id = 'A309695'
system = 'VF06'
passw = 'volvo11'
today=dt.datetime.today()
#todaydsp=today.strftime("%Y%m%d")


dspcutoff=today-dt.timedelta(days=28)
dspcutoff=dspcutoff.strftime("%Y%m%d")
dspcutoff=int(dspcutoff)
week= today.strftime("%Y%W") 
yearweek= str(int(week) + 1)
import win32com.client as win32

dsp= pyodbc.connect('DSN=' + system + ';UID=' + id + ';PWD=' + passw + '')


emails_list=pd.read_excel("//vcn.ds.volvo.net/cli-sd/sd1168/041348/01-Tools/Tools in Progress - Not Validated/EU - AutoReport for RT/Send Reports.xlsm")
emails_list['Dealer']=emails_list['Dealer'].astype(str)

envA=pd.read_sql("SELECT  Trim(BDLNO) as BDLNO, BDSTS, BPPRF, Trim(BPRTN) as BPRTN, BLTIM , BSTBA,BNGSI, BQFC, BQORD, BQBO, BDTLS, BBINN, BPCTXT FROM  G1R3APD.R3F002 WHERE BPPRF='RT' AND BQORD>0  "    , con=dsp )


envB=pd.read_sql("SELECT  Trim(BDLNO) as BDLNO, BDSTS, BPPRF, BLTIM, Trim(BPRTN) as BPRTN, BSTBA, BNGSI, BQFC, BQORD, BQBO, BDTLS, BBINN, BPCTXT FROM G1R3BPD.R3F002 WHERE BPPRF='RT' AND BQORD>0  "   , con=dsp )


dsp = pd.concat([envA,envB])

del envA
del envB
 
dsp=dsp.rename(columns={'BDLNO':'Dealer No.','BLTIM':'Lead Time', 'BDSTS': 'District Number', 'BPRTN': 'Part Number','BPCTXT': 'Purchase Code Text','BPPRF': 'Prefix', 'BQORD': 'On Order in DSP', 'BNGSI':'NEG SI','BQFC': 'Forecast', 'BQBO': 'Quantity on Back Order', 'BSTBA': 'Stock Balance', 'BDTLS':'Date of Last Sale', 'BBINN': 'Binning Location'   }, inplace=False )

dsp= dsp.loc[dsp['Date of Last Sale']<= dspcutoff]

#%%----------Creating a new folder for every new week---------
if not os.path.exists('//vcn.ds.volvo.net/parts-got/proj02/003742/DIM_Analytics_EMEA_and_APAC/07-Process_Automatization/18_RT_Old_on_Order/Files/'+yearweek): 
    os.makedirs('//vcn.ds.volvo.net/parts-got/proj02/003742/DIM_Analytics_EMEA_and_APAC/07-Process_Automatization/18_RT_Old_on_Order/Files/'+yearweek)

#%%----------Comparing DSP data with Extractions from SIPRE and Partslinq and Saving the results in files Dealer wise-------
for dealer in dsp['Dealer No.'].unique():
    temp_dsp=dsp.loc[(dsp['Dealer No.']== dealer)]
    dealer_email=emails_list.loc[emails_list['Dealer']== str(dealer)]
    maxdate=dsp['Lead Time'].max()

    maxdate = math.ceil((maxdate+1) * 7) # Transform that Leadtime in Days
    maxdate = dt.date.today() - dt.timedelta(days=maxdate) # Get Date limit of orders
    
    maxdate= maxdate.strftime("%Y%m%d")
    
    temp_sipre=sipre.loc[(sipre['Dealer No.']== int(dealer))]
    print('Dealer no. ' + str(dealer)+' there are ' + str(temp_dsp.shape[0])+ ' in sipre ' + str(temp_sipre.shape[0]))
    #temp_sipre=temp_sipre.loc[(temp_sipre['date_py']>= int(maxdate))]
    
    print('Dealer no.' + str(dealer)+' there are ' + str(temp_dsp.shape[0])+ ' in sipre ' + str(temp_sipre.shape[0]))
    sum_parts = temp_sipre.filter(items=['Part Number','Quantity on Order', 'Status']) # Only want to see Part and Quantity, the rest is not necessary
    sum_parts = sum_parts.loc[sum_parts['Status'] != 'Cancelled']
    sum_part_shipped= sum_parts
    sum_part_Backorder= sum_parts
    sum_parts= sum_parts.groupby(['Part Number']).agg('sum').reset_index() # Group and Sum per part
    sum_parts['Part Number']=sum_parts['Part Number'].map(str)  
    sum_parts=sum_parts.rename(columns={'Quantity on Order':'Quantity_Confirmed'})
    
    #%%--------------Filter back order information from ware house data------
    sum_part_Backorder= sum_part_Backorder.loc[sum_part_shipped['Status']=='Backorder']
    sum_part_Backorder= sum_part_Backorder.groupby(['Part Number']).agg('sum').reset_index() # Group and Sum per part
    sum_part_Backorder['Part Number']=sum_part_Backorder['Part Number'].map(str) 
    sum_part_Backorder=sum_part_Backorder.rename(columns={'Quantity on Order':'Backorder Qty'})
    
    sum_parts=pd.merge(sum_parts,sum_part_Backorder, how='left')
    
    
    #%%--------- Filtering parts shipped information from ware house data
    
    sum_part_shipped= sum_part_shipped.loc[sum_part_shipped['Status']=='Shipped']
    sum_part_shipped= sum_part_shipped.groupby(['Part Number']).agg('sum').reset_index() # Group and Sum per part
    sum_part_shipped['Part Number']=sum_part_shipped['Part Number'].map(str) 
    sum_part_shipped=sum_part_shipped.rename(columns={'Quantity on Order':'Shipped Qty'})
    
    sum_part_to_be_shipped=pd.merge(sum_parts,sum_part_shipped, how='left')
    
    
    
    #%%------------merge ware house data with DSP data-------
    old_on_order=pd.merge(temp_dsp, sum_part_to_be_shipped, how='left')
    old_on_order['Quantity_Confirmed']=old_on_order['Quantity_Confirmed'].fillna(0)
    old_on_order['Backorder Qty']=old_on_order['Backorder Qty'].fillna(0)
    old_on_order['Shipped Qty']=old_on_order['Shipped Qty'].fillna(0)
    old_on_order['Quantity be Shipped']=old_on_order['Quantity_Confirmed']-old_on_order['Shipped Qty']
    #old_on_order['To be Shipped']=old_on_order['To be Shipped'].fillna(0)
    old_on_order=old_on_order.loc[(old_on_order['On Order in DSP']!=old_on_order['Quantity be Shipped'])]
    old_on_order=old_on_order.loc[ (old_on_order['Quantity be Shipped'] != old_on_order['Backorder Qty']) | (old_on_order['Backorder Qty']==0)]
    del old_on_order['Lead Time'], old_on_order['Quantity on Back Order'] 
    old_on_order=old_on_order[['District Number','Dealer No.', 'Prefix', 'Part Number', 'Purchase Code Text','Date of Last Sale', 'NEG SI', 'Forecast','On Order in DSP','Stock Balance',	'Quantity_Confirmed','Backorder Qty',	'Shipped Qty',	'Quantity be Shipped', 'Binning Location' ]]
    District=str(old_on_order['District Number'].unique())
    
    old_on_order['Action']=''
    old_on_order.loc[(old_on_order['On Order in DSP'] > old_on_order['Quantity be Shipped']), 'Action']= 'Check in DMS'
    old_on_order.loc[(old_on_order['On Order in DSP'] < old_on_order['Quantity be Shipped']), 'Action']= 'Check in SIPRE'
    
    old_on_order=old_on_order.sort_values(by='NEG SI', ascending=False)    
    del old_on_order['Backorder Qty']
    old_on_order['Purchase Code Text'].replace({"FORCT=0 ":"", "LOW-PICK": "", "SLOW    ":"", "FGRPEXCL":""}, inplace=True)
    #%%--------Save file------------------
  
    if len(old_on_order) != 0:
        old_on_order.to_excel('//vcn.ds.volvo.net/parts-got/proj02/003742/DIM_Analytics_EMEA_and_APAC/07-Process_Automatization/18_RT_Old_on_Order/Files/'+yearweek+'/Oldonorder_'+ str(dealer)+'.xlsx', index=False)
        
    if not dealer_email.empty:    
        if  dealer_email.iloc[0,6] == 'X':
    
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = dealer_email.iloc[0,7]
            mail.CC= 'javier.fernandez@renault-trucks.com'
            mail.BCC= 'georges.da.fonte@volvo.com'
            mail.Subject = 'Potential Old on Order to Control' 
            mail.Body = 'Message body'
    
       
            mail.HTMLBody = '''<Font Size = 4 Face= Calibri Color=#2A1E5F>
    
            Dear Dealer,<br><br><br>
            Please find enclosed report for Old on Order <br><br><br>
    
            Kind regards <br>
            
            Muhammad Aarish Jamil<br>
            DIM Analyst<br>
            

            <br><br></font>
            '''      
            mail.Attachments.Add('//vcn.ds.volvo.net/parts-got/proj02/003742/DIM_Analytics_EMEA_and_APAC/07-Process_Automatization/18_RT_Old_on_Order/Files/'+yearweek+'/Oldonorder_'+str(dealer)+'.xlsx')
        
            #mail.Send()
            mail.Display()



   



