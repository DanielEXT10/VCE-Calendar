import pandas as pd
import os
import openpyxl
from pathlib import Path
from tqdm import tqdm
import time

#open("//mx0066inf01/erp_vce$/VCE Calendar/Reports/Reporting summary")
original_directory= os.getcwd()

print(original_directory)

new_directory= r"//mx0066inf01/erp_vce$/VCE Calendar/Reports/Reporting summary"
final_folder=r"\\mx0066inf01\erp_vce$\VCE Calendar"
os.chdir(new_directory)
def concatReports(folder):

    dataframes=[]
    for filename in tqdm(os.listdir()):
        if filename.endswith("xlsx"):
            tempdf = pd.read_excel(open(filename,'rb'), sheet_name="Data Sheet - 1")
            dataframes.append(tempdf)
        time.sleep(0.03)
    return dataframes

reportsList = concatReports(os.getcwd())
df = pd.concat(reportsList)

technicians = pd.read_csv('tecnicians.csv')
ginList= technicians['GIN'].tolist()
ginList = [int(gin) for gin in ginList]
VCEdf = df[df['GIN'].apply(lambda x: x in ginList)]
print(VCEdf)

os.chdir(final_folder)
VCEdf.to_csv("CalendarData.csv")
#VCEdf.to_excel(os.getcwd())