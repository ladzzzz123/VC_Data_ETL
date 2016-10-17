
#new table4
import sqlite3  
import pandas as pd
from pandas import Series, DataFrame
import numpy as np
from xlsxwriter.workbook import Workbook


conn = sqlite3.connect('sample_table4-new.sqlite')
cur = conn.cursor()

cur.execute('''
DROP TABLE IF EXISTS Fund_exit''')

cur.execute('''
CREATE TABLE Fund_exit (org_id TEXT, exit_name TEXT, time TEXT, exit_fund TEXT, exit_type TEXT, exit_return TEXT, unit TEXT)''')

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
            if ('退出分析' in df1.col1.values)==False or ('企业简称' == 
                df1.col1.values.tolist()[df1.col1.values.tolist().index('退出分析')+1])==False:continue                             
     
            org_num=df1.loc[0:0,'col5'].values.tolist()[0] 
            if (org_num in org_col)==True:
                continue
            else:
                org_col.append(org_num)
                
            lst_null=[]
            inds=pd.isnull(df1.col1)
            lst_null=inds[inds==True].index
            start2=int(df1[df1['col1'] == '退出分析'].index[0])+2
            if ('ORG' in org_num)==False:
                print(sheet_name[i],'in','VC',f+1,'机构代码格式错误！')
                continue
                
            if lst_null[-1]>start2:
                end2=lst_null[lst_null>start2][0]-1
                exit_name=df1.loc[start2:end2,'col1'].values.tolist() 
                time=df1.loc[start2:end2,'col2'].values.tolist() 
                exit_fund=df1.loc[start2:end2,'col3'].values.tolist() 
                exit_type=df1.loc[start2:end2,'col4'].values.tolist() 
                exit_return=df1.loc[start2:end2,'col5'].values.tolist() 
                unit =df1.loc[start2:end2,'col6'].values.tolist() 
            else:
                exit_name=df1.loc[start2:,'col1'].values.tolist() 
                time=df1.loc[start2:,'col2'].values.tolist() 
                exit_fund=df1.loc[start2:,'col3'].values.tolist() 
                exit_type=df1.loc[start2:,'col4'].values.tolist() 
                exit_return=df1.loc[start2:,'col5'].values.tolist() 
                unit =df1.loc[start2:,'col6'].values.tolist() 
                
            org_id=[org_num]*len(exit_name)
    
            lstt=[]
            for t in range(len(exit_name)):
                lstt.append((org_id[t],exit_name[t], time[t], exit_fund[t] , exit_type[t] , exit_return[t] , unit[t]))
    
            cur.executemany('''INSERT INTO Fund_exit (org_id, exit_name , time, exit_fund , exit_type , exit_return , unit ) 
            VALUES (?, ?, ?, ?, ?, ?, ? )''', (lstt) )
            
        except:
            print(sheet_name[i],'in VC',f+1,"is problematic")

conn.commit()



#read from .sqlite file and write to xlsx

workbook = Workbook('Table4.xlsx')
worksheet = workbook.add_worksheet()

conn=sqlite3.connect('F://python//sample_table4-new.sqlite')
c=conn.cursor()
c.execute("select * from Fund_exit")
mysel=c.execute("select * from Fund_exit")
for i, row in enumerate(mysel):
    for j, value in enumerate(row):
        worksheet.write(i, j, row[j])
workbook.close()
