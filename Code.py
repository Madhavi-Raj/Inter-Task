#total code
#the required 
#Test parameters: Answered, correct, score, skipped, time-taken, wrong
#Test Names: Concept Test 1, Topic Test 2, Full Chapter Test 1
#are stored in Output.xlsx
#I have used Input_1.xlsx as my test case folder
import pandas as pd
from openpyxl import Workbook
df=pd.read_excel("Input_1.xlsx")
l=['Concept Test 1','Topic Test 2','Full Chapter Test 1']
sh=(df.shape)
rows=sh[0]
columns=sh[1]
cn=list(df.columns)
ans=[]
for i in range(rows):
    rows=list(df.iloc[i])
    if('-' in rows):
        continue
    name=df.iloc[i][0]
    username=df.iloc[i][1]
    chaptertag=df.iloc[i][2]
    j=3
    while(j<columns):
        skip=False
        flag=False
        splitted=cn[j].split(" - ")
        if(splitted[0] in l):
            flag=True
        if(not flag):
            j+=1
        else:
            d=[name,username,chaptertag,splitted[0]]
            k=j
            while(k<j+6):
                flag1=True
                a=(df.iloc[i][k])
                if(pd.isnull(a) or a=='-'):
                    j+=1
                    flag1=False
                    break
                else:
                    d.append(a)
                k+=1
            if(flag1):
                ans.append(d)
                j=k
            else:
                skip=True
        if(skip):
            break
wb = Workbook() # creates a workbook object.
ws = wb.active # creates a worksheet object.
ws.append(['Name','Username','Chapter Tag','Test_Name','answered','correct','score','skipped','time-taken (seconds)','wrong'])
for l in ans:
    my_r=[l[0],l[1],l[2],l[3],l[6],l[7],l[4],l[9],l[5],l[8]]
    ws.append(my_r)
wb.save('Output.xlsx')
