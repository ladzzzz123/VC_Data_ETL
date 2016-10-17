
#new table1
import sqlite3  
import pandas as pd
from pandas import Series, DataFrame
import numpy as np


conn = sqlite3.connect('sample_table1-new.sqlite')
cur = conn.cursor()

cur.execute('''
DROP TABLE IF EXISTS Org_Info''')

cur.execute('''
CREATE TABLE Org_Info (org_id TEXT, org_name TEXT, org_short TEXT, eng_name TEXT, found_time TEXT, org_type TEXT, capital_source TEXT, 
HQ TEXT, website TEXT, former_name TEXT, close_time TEXT, office TEXT, description TEXT )''')

filename=[]
org_col=[]
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
 
            org_id=df1.loc[0:0,'col5'].values.tolist()[0] 
            if (org_id in org_col)==True:
                continue
            else:
                org_col.append(org_id)
                
            org_name=df1.loc[0:0,'col1'].values.tolist()[0] 
            eng_name=df1.loc[1:1,'col1'].values.tolist()[0] 
            found_time=df1.loc[6:6,'col2'].values.tolist()[0] 
            org_type=df1.loc[7:7,'col2'].values.tolist()[0] 
            capital_source=df1.loc[8:8,'col2'].values.tolist()[0] 
            HQ=df1.loc[9:9,'col2'].values.tolist()[0] 
            website=df1.loc[6:6,'col5'].values.tolist()[0] 
            former_name=df1.loc[7:7,'col5'].values.tolist()[0] 
            close_time=df1.loc[8:8,'col5'].values.tolist()[0] 
            office=df1.loc[9:9,'col5'].values.tolist()[0] 
            org_short=df1.loc[4:4,'col2'].values.tolist()[0] 
            if ('ORG' in org_id)==False:
                print(sheet_name[i],'in','VC',f+1,'投资机构代码错误！')
                continue
            #description=
            pos=df1[df1['col1'] == '机构描述'].index[0]+1
            description=df1.loc[pos:pos,'col1'].values.tolist()[0]
    
            cur.execute('''INSERT INTO Org_Info (org_id, org_name, org_short, eng_name, found_time, org_type,capital_source, 
            HQ, website, former_name, close_time, office, description) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',(org_id, org_name,org_short, eng_name, found_time, org_type, capital_source, HQ,
                                        website,former_name, close_time, office, description) )
            
        except:
            print(sheet_name[i],'in VC',f+1,"is problematic")

conn.commit()
