# -*- coding: utf-8 -*-
"""
Created on Wed Jun 10 11:02:19 2020
@author: tao_jingfen
"""
import os
import pandas as pd
from collections import defaultdict
import time

def excel2Dict(excelPath : str):
    """
    read excel data into dict
    ::input excel file path 
    excel with three sheets
    -------------------------------------------------------------------------
    sheet1: protein each purification step infomation
    sheet1 title: 
    [ ProjectName	
    Date
    PurificationStepNo
    PurificationStep
    ProteinNo	
    ProteinName
    Concentration	
    Volume	
    Amount
    Yield
    Buffer
    Purity by SEC-HPLC(%)
    Recovery(%)
    MW
    PI
    Supernatant
    Comments ]
    sheet2: each step SDS and HPLC information
    sheet2 title:
    [ PurificationStepNo
    PurificationStep
    SDS_Picture	
    SDS_Table
    SDS_ELN
    SDS_Conclusion
    HPLC_Picture
    HPLC_ELN
    HPLC_Conclusion ]
    ------------------------------------------------------------------------
    sheet2: SDS-PAGE table
    sheet2 title:
    [ Table
    Lane
    Protein name
    MW(kDa) ]
    ::output proteinInfoDict,sdsTableDict
    -------------------------------------------------------------------------
    1. proteinInfoDict: 
    {
        'ProjectName': 'WBPXXX','Date': '04/13/2020','Purification': 
        {
            1:{
            'ProteinName':,
            'PurificationStepsType':,
            1: {'PurificationStep': 'Protein A',
            'Concentration':,
            ...,
            'HPLC_Conclusion':},
            'Dialysis':{}
            },
            2:{},3:{},4:{}                                               }
        }
    }
    -------------------------------------------------------------------------
    2. sdsTableDict: 
    {
        'SupernatantDialysis': {'Lane': [],'Protein name': [],'MW(kDa)': []},
        'SEC':                 {'Lane': [],'Protein name': [],'MW(kDa)': []},
        'CEX':                 {'Lane': [],'Protein name': [],'MW(kDa)': []}
    }
    --------------------------------------------------------------------------
    """
    # -------------------read sheet1 to proteinInfoDict-----------------------
    sheet1 = pd.read_excel(excelPath,keep_default_na=False,dtype=str)
    for col in ['Concentration (mg/ml)','Volume (ml)','Amount (mg)',
                'Yield (mg/L)','PI']:
        for index in sheet1.index:
            if sheet1[col][index] != '':
                sheet1[col][index] = pd.to_numeric(sheet1[col][index])
                sheet1[col][index] = ('%.2f')% (sheet1[col][index])
    
    ### coverPage
    if len(sheet1['ProjectName']) != 0:
        projectName = sheet1['ProjectName'][0]
        date = sheet1['Date'][0]
    else:
        projectName = 'project'
        date = time.strftime("%m/%d/%Y", time.localtime())
    
    ### finalPage
    finalDict = defaultdict(dict)
    if len(sheet1['PurificationStepNo']) != 0:
        df1 = sheet1.iloc[:,2:]
        df_group = df1[['PurificationStepNo','ProteinName']].groupby(by='ProteinName',as_index=False).max()
        df_merge = pd.merge(df_group,df1,on=['PurificationStepNo','ProteinName'],how='left')
        df_merge = df_merge.sort_values(by='ProteinNo')
        for index in df_merge.index:
            proteinName = df_merge['ProteinName'][index]
            finalDict[proteinName]['No'] = str(df_merge['ProteinNo'][index])
            for col in df_merge.columns.tolist()[4:]:
                colName = col.split("(")[0].strip()
                finalDict[proteinName][colName] = str(df_merge[col][index])
        
    ### processPage
    processList = []
    if len(sheet1['ProteinNo']) != 0:
        processDict = defaultdict(list)
        for i in df1['ProteinNo'].drop_duplicates():
            a = list(df1[['ProteinNo','PurificationStep']][df1['ProteinNo'] == i]['PurificationStep'])
            PurificationStepOnly =  [a[0]]+[a[i] for i in range(1,len(a)) if a[i] != a[i-1]]
            process = '_'.join(PurificationStepOnly)
            processDict[process].append(i)
        for key,value in processDict.items():
            oneProcessList = []
            oneProcessList.append('Protein # '+ ', '.join(value))
            oneProcessList.extend(key.split('_'))
            oneProcessList.append('Filtration & Storage')
            processList.append(oneProcessList) 

    ### stepPage
    stepsDict = defaultdict(dict)
    if len(sheet1['PurificationStepNo']) != 0:
        for i in df1['PurificationStepNo'].drop_duplicates():
            tempDf = df1[df1['PurificationStepNo'] == i]
            step = ' & '.join(tempDf['PurificationStep'].drop_duplicates())
            stepsDict[i]['PurificationStep'] = step
            for index in tempDf.index:
                proteinName = tempDf['ProteinName'][index]
                stepsDict[i][proteinName] = {}
                stepsDict[i][proteinName]['No'] = str(tempDf['ProteinNo'][index])
                for col in tempDf.columns.tolist()[4:]:
                    colName = col.split("(")[0].strip()
                    stepsDict[i][proteinName][colName] = str(tempDf[col][index])

    ### sdsPage
    if len(sheet1['Supernatant (mL)']) != 0:
        supernatant = str(df1['Supernatant (mL)'][0])
    else:
        supernatant = '0'
    sdsList = []
    sdsTableDict = defaultdict(dict)
    sheet2 = pd.read_excel(excelPath,sheet_name=1,keep_default_na=False,dtype=str)
    sheet3 = pd.read_excel(excelPath,sheet_name=2,keep_default_na=False,dtype=str)
    for index in sheet2.index:
        sdsStepList = []
        for col in sheet2.columns.tolist()[1:6]:
            sdsStepList.append(sheet2[col][index])
        sdsList.append(sdsStepList)
    for index in sheet3.index:
        for col in sheet3.columns[1:]:
            if col not in sdsTableDict[sheet3['Table'][index]].keys():
                sdsTableDict[sheet3['Table'][index]][col] = [sheet3[col][index]]
            else:  
                sdsTableDict[sheet3['Table'][index]][col].append(sheet3[col][index]) 
    
    ### hplcPage                
    hplcList = []
    proDict = defaultdict(dict)
    for index in sheet2.index:
        hplcStepList = []
        for col in sheet2.columns.tolist()[0:2] + sheet2.columns.tolist()[6:]:
            hplcStepList.append(sheet2[col][index])
        hplcList.append(hplcStepList)  
    if len(sheet1['PurificationStepNo']) != 0:
        df1 = sheet1.iloc[:,2:]                     
        for i in df1['PurificationStepNo'].drop_duplicates():
            tempDf = df1[df1['PurificationStepNo'] == i]
            step = ' & '.join(tempDf['PurificationStep'].drop_duplicates())
            step = '_'.join((i, step.replace(' ','_')))
            proDict[step] = {}
            for index in tempDf.index:
                proteinName = tempDf['ProteinName'][index]
                proDict[step][proteinName] = [str(tempDf['ProteinNo'][index])]
                proDict[step][proteinName].append(tempDf['Purity by SEC-HPLC (%)'][index])

    return (projectName,
            date),finalDict,processList,stepsDict,(supernatant,
                sdsList,sdsTableDict),hplcList,proDict

if __name__ == '__main__':
    excelPath = 'D:\\2.PS work\\2020\\3.auto report\\example\\Demo2\\demo2.xlsx'
    (projectName,date),finalDict,processList,stepsDict,(supernatant,sdsList,
    sdsTableDict),hplcList,proDict = excel2Dict(excelPath)
    print(projectName)
    print(date)
    print(finalDict)
    print(processList)
    print(stepsDict)
    print(supernatant)
    print(sdsList)
    print(sdsTableDict)
    print(hplcList)
    print(proDict)