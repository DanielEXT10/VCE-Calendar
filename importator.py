
import pandas as pd
import os
import openpyxl
from pathlib import Path
from tqdm import tqdm
import time


    
folderpath = (r"C:\Users\DPerez48\Documents\Excel Projects\LC Calendar\Reports\Reporting summary")
#folderpath = ("\\mx0066inf01\erp_vce$\VCE Calendar\Reports\Reporting summary")
def concatReports(folder):
    dataframes=[]
    for filename in tqdm(os.listdir(folderpath)):
        if filename.endswith(".xlsx"):
            tempdf = pd.read_excel(open(filename,'rb'), sheet_name="Data Sheet - 1")
            dataframes.append(tempdf)
        time.sleep(0.03)
    return dataframes
    
reportsList =concatReports(folderpath)
df = pd.concat(reportsList)

technicians = pd.read_csv('tecnicians.csv')
ginList= technicians['GIN'].tolist()
ginList = [int(gin) for gin in ginList]
VCEdf = df[df['GIN'].apply(lambda x: x in ginList)]
VCEdf.to_excel(r'C:\Users\DPerez48\Documents\Excel Projects\LC Calendar\VCEdata.xlsx')
#VCEdf.to_excel("\\mx0066inf01\erp_vce$\VCE Calendar\Reports")


