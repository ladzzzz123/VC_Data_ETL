import pandas as pd

def read_file(path):
    df_all = pd.read_excel(path, header=0) #Y://Data//python//Liq//Table2_new.xlsx
    return df_all                          #header=0, i.e. first row parsed as column name

    
#reason: pd read year data as float, thus need to change
def transformDF(df_all):
    oldtime1 = df_all['投资时间'].tolist()
    newtime1 = []
    for i in range(len(oldtime1)):
        if '.' not in str(oldtime1[i]):
            newtime1.append('')
        else:
            newtime1.append(str(int(oldtime1[i])))

    # replace this col
    df_all_new = df_all.replace(to_replace=df_all['投资时间'].tolist(), value=newtime1)
    return df_all_new

#common function, unique elements with a same order 
def unilst(lst):
    uni=[]
    for i in range(len(lst)):
        if lst[i] in uni:
            continue
        else:
            uni.append(lst[i])
    return uni

def mainprocess(df_all_new):
    org = df_all_new['机构代码'].tolist()
    uni_org = unilst(org)
    t2_orgnum = [] #(org,year) as primary key
    t2_year = []
    t2_exp = [] #investment experience (number of companies invested/year)
    t2_idscul = []#cumulative number of industries this org has entered
    t2_vca = []  #type VC-series A
    t2_vcb = []
    t2_vcc = []
    t2_vcd = []
    t2_vce = []
    t2_vcm = [] # missing value count
    t2_newids = [] #number of new industries entered/year
    t2_newdist = [] #number of new districts entered/year
    for i in range(len(uni_org)):  #loop of orgs
        df_suborg = df_all_new[df_all_new['机构代码'] == uni_org[i]]

        lst_tempyear = [int(ele) for ele in df_suborg['投资时间'] if ele != '']
        if len(lst_tempyear) == 0:
            continue

        max_year = max(lst_tempyear)
        min_year = min(lst_tempyear)

        industry_cum = 0
        industry_lst = []
        new_district_lst = []
        for j in range(max_year - min_year + 1): #loop of years /org
            df_subyear = df_suborg[df_suborg['投资时间'] == str(min_year + j)]
            VC_A = 0
            VC_B = 0
            VC_C = 0
            VC_D = 0
            VC_E = 0
            VC_miss = 0
            new_industry = 0
            new_district = 0
            #print(min_year + j)
            #print('length', len(df_subyear))
            #print(df_subyear)

            if len(df_subyear) == 0: #no data here
                year = str(min_year + j)
                exp = 0
                industry_cum = industry_cum
                VC_A = 0
                VC_B = 0
                VC_C = 0
                VC_D = 0
                VC_E = 0
                VC_miss = 0
                new_industry = 0
                new_district = 0
            else:
                year = str(min_year + j)
                exp = len(df_subyear['投资企业'])

                for k in range(len(df_subyear['行业'])):
                    #print(df_subyear['行业'].tolist()[k])
                    if df_subyear['行业'].tolist()[k] == '':
                        continue
                    if df_subyear['行业'].tolist()[k] not in industry_lst:
                        industry_lst.append(df_subyear['行业'].tolist()[k])
                        new_industry += 1
                        industry_cum += 1
                    else:
                        continue

                for k in range(len(df_subyear['投资性质'])):
                    if df_subyear['投资性质'].tolist()[k] == '':
                        VC_miss += 1
                        continue
                    if df_subyear['投资性质'].tolist()[k] == 'VC-Series A':
                        VC_A += 1
                    elif df_subyear['投资性质'].tolist()[k] == 'VC-Series B':
                        VC_B += 1
                    elif df_subyear['投资性质'].tolist()[k] == 'VC-Series C':
                        VC_C += 1
                    elif df_subyear['投资性质'].tolist()[k] == 'VC-Series D':
                        VC_D += 1
                    elif df_subyear['投资性质'].tolist()[k] == 'VC-Series E':
                        VC_E += 1
                    else:
                        VC_miss += 1

                for k in range(len(df_subyear['地区'])):
                    if df_subyear['地区'].tolist()[k] == '':
                        continue
                    if df_subyear['地区'].tolist()[k] not in new_district_lst:
                        new_district += 1
                        new_district_lst.append(df_subyear['地区'].tolist()[k])
                    else:
                        continue
            # print(industry_cum)
            t2_orgnum.append(uni_org[i])
            t2_year.append(year)
            t2_exp.append(exp)
            t2_idscul.append(industry_cum)
            t2_vca.append(VC_A)
            t2_vcb.append(VC_B)
            t2_vcc.append(VC_C)
            t2_vcd.append(VC_D)
            t2_vce.append(VC_E)
            t2_vcm.append(VC_miss)
            t2_newids.append(new_industry)
            t2_newdist.append(new_district)

            print("org:%10s,year:%5s,exp:%d,induscul:%d,A:%d,B:%d,C:%d,D:%d,E:%d,Miss:%d,newindus:%d,newdist:%d" %
                  (uni_org[i], year, exp, industry_cum, VC_A, VC_B, VC_C, VC_D, VC_E, VC_miss, new_industry, new_district))



#multiple lists into a single dataframe
def lst_to_df(t2_orgnum,t2_year,t2_exp,t2_idscul,t2_vca,t2_vcb,t2_vcc,t2_vcd,t2_vce,t2_vcm,t2_newids,t2_newdist):
    t2df = pd.DataFrame(
        {'ORG_NUM': t2_orgnum,
         'YEAR': t2_year,
         'EXPERIENCE/year': t2_exp,
         'CUMUL-INDUSTRY': t2_idscul,
         'VC_typeA': t2_vca,
         'VC_typeB': t2_vcb,
         'VC_typeC': t2_vcc,
         'VC_typeD': t2_vcd,
         'VC_typeE': t2_vce,
         'VC_miss': t2_vcm,
         'NEW_INDUSTRY/year': t2_newids,
         'NEW_DISTRICT/year': t2_newdist
         })
    #rearrange the order of columns     
    t2df_f = t2df[
        ['ORG_NUM', 'YEAR', 'EXPERIENCE/year', 'CUMUL-INDUSTRY', 'VC_typeA', 'VC_typeB', 'VC_typeC', 'VC_typeD',
         'VC_typeE',
         'VC_miss', 'NEW_INDUSTRY/year', 'NEW_DISTRICT/year']]
    return t2df_f

#write a df to a xlsx file
def df_to_xlsx(df,path):
    writer = pd.ExcelWriter(path)
    df.to_excel(writer, 'Sheet1')
    writer.save()

def check_list_value():
    disset=set(df_all['投资性质'].tolist())
    print(disset)


def main():
    df_all=read_file('Y://Data//python//Liq//Table2_new.xlsx')
    df_all_new=transformDF(df_all)
    mainprocess(df_all_new)


if __name__ == '__main__':
    main()

