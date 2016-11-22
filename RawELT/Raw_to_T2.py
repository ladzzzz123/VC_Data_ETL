
#new table2
import sqlite3  
import pandas as pd
from pandas import Series, DataFrame
import numpy as np
conn = sqlite3.connect('sample_table2-new.sqlite')
cur = conn.cursor()

cur.execute('''
DROP TABLE IF EXISTS Investment''')

cur.execute('''
CREATE TABLE Investment (org_id TEXT, company TEXT, industry TEXT, district TEXT, 
investor TEXT, time TEXT, inv_property TEXT)''')

filename=[]
for k in range(6):
    filename.append('F://python//VC%s'%(k+1)+'.xls') #initialize all the files

org_col=[]
for f in range(6):
    xls = pd.ExcelFile(filename[f])
    print('file No.',f+1)
    sheet_len=len(xls.sheet_names)
    sheet_name=[]
    
    for j in range(sheet_len):
        sheet_name.append("Sheet%s"%(j+1))    #iterate through all the sheets
    for i in range(sheet_len):
        try:                   #powerful try & except!
            df1 = xls.parse(sheet_name[i])
            df1.columns = ['col1', 'col2','col3','col4','col5','col6', 'col7']  #rename columns
            #second table
            if ('投资事件' in df1.col1.values)==False or ('企业简称' == 
                                                      df1.col1.values.tolist()[df1.col1.values.tolist().index('投资事件')+1])==False:continue                             
     
            org_num=df1.loc[0:0,'col5'].values.tolist()[0] 
            if (org_num in org_col)==True:
                continue
            else:
                org_col.append(org_num)
            
            lst_null=[]
            inds=pd.isnull(df1.col1)
            lst_null=inds[inds==True].index
            start2=int(df1[df1['col1'] == '企业简称'].index[0])+1
            if ('ORG' in org_num)==False:
                print(sheet_name[i],'in','VC',f+1,'机构代码格式错误！')
                continue
             
            if lst_null[-1]>start2:
                end2=lst_null[lst_null>start2][0]-1
                industry=df1.loc[start2:end2,'col2'].values.tolist() 
                company=df1.loc[start2:end2,'col1'].values.tolist() 
                district=df1.loc[start2:end2,'col3'].values.tolist() 
                investor=df1.loc[start2:end2,'col4'].values.tolist() 
                time=df1.loc[start2:end2,'col5'].values.tolist() 
                inv_property=df1.loc[start2:end2,'col6'].values.tolist() 
            else:
                industry=df1.loc[start2:,'col2'].values.tolist() 
                company=df1.loc[start2:,'col1'].values.tolist() 
                district=df1.loc[start2:,'col3'].values.tolist() 
                investor=df1.loc[start2:,'col4'].values.tolist() 
                time=df1.loc[start2:,'col5'].values.tolist() 
                inv_property=df1.loc[start2:,'col6'].values.tolist() 
            org_id=[org_num]*len(company)
    
            lstt=[]
            for t in range(len(company)):
                lstt.append((org_id[t],company[t],industry[t], district[t], investor[t], time[t], inv_property[t]))
    
            cur.executemany('''INSERT INTO Investment (org_id, company, industry,district, investor, time, inv_property) 
            VALUES (?, ?, ?, ?, ?, ?, ? )''', (lstt) )
            
        except:
            print(sheet_name[i],'in VC',f+1,"is problematic")

conn.commit()

