import pandas as pd
import os
import json


def monthlyCompliance():
    fileName = str(os.path.dirname(os.path.realpath(__file__))) + '\\data\\' + 'DISCO DISPATCH BENCHMARK' + '.xlsx'
    sheet = 'eko'
    df = pd.read_excel(io=fileName,sheetname=sheet)
    fileName = str(os.path.dirname(os.path.realpath(__file__))) + '\\data\\disco\\' + sheet + '-data\\' + 'otherData' +'.json' 
    json_data = str(round((float(df.iloc[105,100]) * 100),1)) + '%'
    with open(fileName, 'w') as fp:
        json.dump(json_data,fp,sort_keys=True, indent = 4)
    #print(df.head(119))
    
def determineFunc(day):
    fileName = str(os.path.dirname(os.path.realpath(__file__))) + '\\data\\' + 'DISCO DISPATCH BENCHMARK' + '.xlsx'
    sheet = 'eko'
    df = pd.read_excel(io=fileName,sheetname=sheet)
    #print(df.iloc[:,4])
    discoNames = []
    compliance =[]
    print('running: ',day)
    for a in range(5,(len(df.iloc[:,4]))-5):
        discoNames.append(df.iloc[a,4])
    if (day != 1):
        control = (9+(3*(day-1)))
        #print(df.iloc[:,control])
        for b in range(5,(len(df.iloc[:,control])-5)):
            if(str(df.iloc[b,control]) == str('nan')):
                compliance.append('')
            else:
                compliance.append(df.iloc[b,control])
           
            
            
    else:
        df.iloc[:,9]
        for c in range(5,(len(df.iloc[:,9])-5)):
            if(str(df.iloc[c,9]) == str('nan')):
                compliance.append('')
            else:
                compliance.append(df.iloc[c,9])
      
    fileName = str(os.path.dirname(os.path.realpath(__file__))) + '\\data\\disco\\' + sheet + '-data\\' + str(day) +'.json' 
    data = dict(zip(discoNames,compliance))
    json_data = data


    with open(fileName, 'w') as fp:
        json.dump(json_data,fp,sort_keys=True, indent = 4)

for x in range(1,32):
    monthlyCompliance()
    determineFunc(x)

""" os.system('c:\python34\python -m http.server')
os.system("c:\python34\pythonpython -m webbrowser -t 'http://localhost:8000/nerc.html'")

 """