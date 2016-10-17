#Nice Code
#new table3
import sqlite3  
import pandas as pd
from pandas import Series, DataFrame
import numpy as np
conn = sqlite3.connect('sample_table3-new.sqlite')
cur = conn.cursor()

cur.execute('''
DROP TABLE IF EXISTS Cooperation''')

cur.execute('''
CREATE TABLE Cooperation (org_id TEXT, coopetation_comp TEXT, time TEXT, object_invest TEXT)''')

org_col=[]
filename=[]
for k in range(6):
    filename.append('F://python//VC%s'%(k+1)+'.xls') #initialize all the files

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
            if ('合作关系' in df1.col1.values)==False or ('合作投资机构' == 
                                                      df1.col1.values.tolist()[df1.col1.values.tolist().index('合作关系')+1])==False:continue                             
     
            org_num=df1.loc[0:0,'col5'].values.tolist()[0] 
            if (org_num in org_col)==True:
                continue
            else:
                org_col.append(org_num)
            lst_null=[]
            inds=pd.isnull(df1.col1)
            lst_null=inds[inds==True].index
            start2=int(df1[df1['col1'] == '合作投资机构'].index[0])+1
            if ('ORG' in org_num)==False:
                print(sheet_name[i],'in','VC',f+1,'机构代码格式错误！')
                continue
             
            if lst_null[-1]>start2:
                end2=lst_null[lst_null>start2][0]-1
                coopetation_comp=df1.loc[start2:end2,'col1'].values.tolist() 
                time=df1.loc[start2:end2,'col2'].values.tolist() 
                object_invest=df1.loc[start2:end2,'col3'].values.tolist() 
                
            else:
                coopetation_comp=df1.loc[start2:,'col1'].values.tolist() 
                time=df1.loc[start2:,'col2'].values.tolist() 
                object_invest=df1.loc[start2:,'col3'].values.tolist() 
                
            org_id=[org_num]*len(coopetation_comp)
    
            lstt=[]
            for t in range(len(coopetation_comp)):
                lstt.append((org_id[t],coopetation_comp[t], time[t], object_invest[t]))
    
            cur.executemany('''INSERT INTO Cooperation (org_id, coopetation_comp, time, object_invest) 
            VALUES (?, ?, ?, ? )''', (lstt) )
            
        except:
            print(sheet_name[i],'in VC',f+1,"is problematic")

conn.commit()
