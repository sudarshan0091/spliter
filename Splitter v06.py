#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd 
import os 
import datetime as dt
import warnings
warnings.filterwarnings("ignore", 'This pattern has match groups')


# In[2]:


col_list = ['Company Name',  'Airwaybill','Invoice', 'Pickup Date', 'Delivery Date', 'PDD','Pickup City', 
            'Ship City', 'Ship Pin Code', 'No of boxes','Actual Weight', 'Total Attempt Count', 'XB Status','Remarks',]
df = pd.read_excel(input("Enter file path which you want to split "), usecols=col_list)
column_name = "Company Name"
replace_symbols =[">","<",":","/", "\\\\","\|", "\?", "\*","(", ")"]
df[column_name] = df[column_name].replace(replace_symbols).str.strip().str.title()
unique_values = df[column_name].unique()


df['Pickup Date'] = df['Pickup Date'].dt.date
df['Delivery Date'] = df['Delivery Date'].dt.date
df['PDD'] = df['PDD'].dt.date


# In[3]:


for unique_value in unique_values:
    df_output = df[df[column_name].str.contains(unique_value)]
    output_path = os.path.join(r"C:\Users\sudarshan\Desktop\output", str(unique_value) + '.xlsx')
    #df_output.to_excel(output_path, sheet_name=unique_value[:31], index=False,columns=['Company Name', 'Airwaybill', 'Pickup Date', 'Delivery Date', 'PDD','Route Mode', 'Pickup City', 'Ship City', 'Ship Pin Code', 'No of boxes','Actual Weight', 'Total Attempt Count', 'XB Status','Remarks'])
    df_output.to_excel(output_path, sheet_name=unique_value[:31], index=False)


# In[ ]:



input()

