# -*- coding: utf-8 -*-
"""
Created on Mon Jan 20 12:36:08 2020

@author: sbiswas149 - sankar biswas
"""

import csv
import xlrd
import re
import os
import xlsxwriter
import pandas as pd
import openpyxl 
from pyxlsb import open_workbook
import numpy as np

outputCol = 0
outputRow = 1

os.chdir(r"C:\Users\sb512911\Desktop\All\Applications\VDR\output")
workbook = xlsxwriter.Workbook('demo.xlsx')

#sheet - Raw Details
worksheet1 = workbook.add_worksheet("Raw Details")
bold = workbook.add_format({'bold': True})
worksheet1.write('A1','File Name', bold)
worksheet1.write('B1','Index Number', bold)
worksheet1.write('C1','Folder Name', bold)
worksheet1.write('D1','File Path', bold)
worksheet1.write('E1','File Type', bold)
worksheet1.write('F1','Total Number of Page(s)/Slide(s)', bold)
worksheet1.write('G1','Keyword', bold)
worksheet1.write('H1','IT Category', bold)
worksheet1.write('I1','Page/Slide Number', bold)

#sheet - File Level Detail
"""
worksheet2 = workbook.add_worksheet("File Level Detail")
worksheet2.write('A1','File Name', bold)
worksheet2.write('B1','Index Number', bold)
worksheet2.write('C1','Folder Name', bold)
worksheet2.write('E1','File Type', bold)
worksheet2.write('E1','Scan Status', bold)
worksheet2.write('F1','Number of Keywords found', bold)
worksheet2.write('F1','Number of Keyword Categories', bold)
worksheet2.write('G1','Keyword Details', bold)
worksheet2.write('H1','Category Details', bold)
worksheet2.write('D1','File Path', bold)
worksheet2.write('I1','Relevant?', bold)
worksheet2.write('I1','Applicable DRL Items', bold)

#sheet - Keyword Details
worksheet3 = workbook.add_worksheet("Keyword Details")
worksheet3.write('A1','Keyword', bold)
worksheet3.write('B1','Keyword Category', bold)
worksheet3.write('C1','Number of files with keywords', bold)
worksheet3.write('E1','Total keyword hits', bold)
worksheet3.write('E1','File(s) with maximum keyword hits', bold)
worksheet3.write('F1','File Path(s)', bold)
worksheet3.write('F1','All file(s) with keyword hits', bold)
worksheet3.write('G1','Index Number(s)', bold)
worksheet3.write('H1','File Path(s)', bold)
"""
#sheet - DRL Details
worksheet4 = workbook.add_worksheet("DRL Details")
worksheet4.write('A1','DRL #', bold)
worksheet4.write('B1','Request', bold)
worksheet4.write('C1','Relevant Files Found?', bold)
worksheet4.write('D1','Number of relevant files', bold)
worksheet4.write('E1','File Name(s)', bold)
worksheet4.write('F1','Index Number(s)', bold)
worksheet4.write('G1','File Path(s)', bold)


def specialCharReplace(word):
    return re.sub('[^A-Za-z0-9\s]+', '', word)

def cell_header(a):
    if a <= 26:
        return chr(a+64)
    count = a//26 + 64
    rest = a%26 + 64
    return chr(count) + chr(rest)

def csv_handler(path):
    with open(path) as csvfile:
        readCSV = csv.reader(csvfile)
        count = 0
        page_count = 1
        for row in readCSV:
            # count is row number 
            count = count + 1
            for i in range(len(row)):
                #print(row[i] + " " + str(count) + " " + str(i))
                # row is representing each row  of the csv file - a list 
                list_details = folderName(path)
                cell_details = cell_header(i+1)+str(count)
                sheet_handler(list_details, cell_details, path, page_count, row[i])
                
def outputWriterRawDetail(file_name, index_number, folder_name, file_path, file_type, pages, keyword, category, cell):
    global outputRow
    global outputCol
    worksheet1.write(outputRow,0,file_name)
    worksheet1.write(outputRow,1,index_number)
    worksheet1.write(outputRow,2,folder_name)
    worksheet1.write(outputRow,3,file_path)
    worksheet1.write(outputRow,4,file_type)
    worksheet1.write(outputRow,5,pages)
    worksheet1.write(outputRow,6,keyword)
    worksheet1.write(outputRow,7,category)
    worksheet1.write(outputRow,8,cell)
    outputRow = outputRow + 1
    
def sheet_handler(list_details, cell_details, path, page_count, row):
    if isinstance(row, str):
        mm1 = [i for i in r1.findall(row) if i and isinstance(i, str)]
        if mm1:
            for keyword in mm1:
                outputWriterRawDetail(list_details[0], list_details[1], list_details[2], path, list_details[3], page_count,
                        keyword, "IT Spend", cell_details)
        mm2 = [i for i in r2.findall(row) if i and isinstance(i, str)]                   
        if mm2:
            for keyword in mm2:
                outputWriterRawDetail(list_details[0], list_details[1], list_details[2], path, list_details[3], page_count,
                        keyword, "IT Organization / Roles", cell_details)
        """
        mm3 = r3.findall(row[i])
        if mm3:
            list_details = folderName(path)
            cell_details = cell_header(i+1)+str(count)
            outputWriterRawDetail(list_details[0], list_details[1], list_details[2], path, list_details[3], 1,
                    row[i], "IT Applications", cell_details)
        """
        mm4 = [i for i in r4.findall(row) if i and isinstance(i, str)]
        if mm4:
            for keyword in mm4:
                print(keyword)
                outputWriterRawDetail(list_details[0], list_details[1], list_details[2], path, list_details[3], page_count,
                        keyword, "IT Infrastructure", cell_details)
                        
        mm5 = [i for i in r5.findall(row) if i and isinstance(i, str)]
        if mm5:
            for keyword in mm5:
                outputWriterRawDetail(list_details[0], list_details[1], list_details[2], path, list_details[3], page_count,
                        keyword, "IT Security and Controls", cell_details)
                        
        mm6 = [i for i in r6.findall(row) if i and isinstance(i, str)]
        if mm6:
            for keyword in mm6:
                outputWriterRawDetail(list_details[0], list_details[1], list_details[2], path, list_details[3], page_count,
                        keyword, "IT Projects", cell_details)
                        
        mm7 = [i for i in r7.findall(row) if i and isinstance(i, str)]
        if mm7:
            for keyword in mm7:
                outputWriterRawDetail(list_details[0], list_details[1], list_details[2], path, list_details[3], page_count,
                        keyword, "Target Technology Fitness Scorecard", cell_details)
                        
        mm8 = [i for i in r8.findall(row) if i and isinstance(i, str)]
        if mm8:
            for keyword in mm8:
                outputWriterRawDetail(list_details[0], list_details[1], list_details[2], path, list_details[3], page_count,
                        keyword, "Misecellaneous", cell_details)         
    
    
  
def xlsx_handler(path):
    sheets_dict = pd.read_excel(path, sheet_name=None, header = None)
    sheet_count = len(sheets_dict)
    for name, sheet in sheets_dict.items():
        for index, row in sheet.iterrows():
            for i in range(len(row)):
                list_details = folderName(path)
                cell = cell_header(i+1)+str(index+1)
                cell_details = name + ": " + cell
                sheet_handler(list_details, cell_details, path, sheet_count, row[i])
                
def xlsb_handler(path): 
    sheet_count = len(open_workbook(path).sheets)
    with open_workbook(path) as wb:
        for sheetname in wb.sheets:
            row_number = 0
            with wb.get_sheet(sheetname) as sheet:
                for row in sheet.rows():
                    row_number = row_number + 1
                    for i in range(len(row)):
                        if row[i].v:
                            list_details = folderName(path)
                            cell = cell_header(i+1)+str(row_number)
                            cell_details = sheetname + ": " + cell
                            sheet_handler(list_details, cell_details, path, sheet_count, row[i].v)      
   
def folderName(path):
    #C:\Users\sb512911\Desktop\All\Screenshots\1.2.1 ILS Structure Chart.csv
    file_name = path.split("\\")[-1]
    index_number = file_name.split(" ")[0]
    folder_name = path.split(file_name)[0]
    file_type = path.split(".")[-1]   
    return [file_name, index_number, folder_name, file_type]

# pass the additional parameter for relevance test - dynamically
def generatingFileLevelDetail(file_path):      
    data = pd.read_excel(open(file_path,'rb'), sheet_name=0)
    if not data.empty:
    #writing to "File Level Detail"
        file_level_detail = data[["File Name","Index Number", "Folder Name", "File Type"]].drop_duplicates()
        file_level_detail["Number of Keywords found"] = \
                data[['File Name']].groupby(['File Name']).size().reset_index()[0].values
        
        file_level_detail["Number of Keyword Categories"] = \
                data[['File Name', 'IT Category']].drop_duplicates().groupby(['File Name']).size()\
                .reset_index()[0].values
                
        intermediate1 = data[['File Name','Keyword']].groupby(['File Name', 'Keyword']).size().reset_index()
        intermediate1[[0]] = intermediate1[[0]].astype(str).values
        intermediate1['joined'] = intermediate1[['Keyword', 0]].apply(lambda x: '('.join(x)+"), ", axis =1).values
                
        file_level_detail["Keyword Details"] = intermediate1[["File Name", "joined"]].groupby(["File Name"]).sum().values
        
        file_level_detail["Category Details"] = \
            data.groupby("File Name").apply(lambda x : x["IT Category"].drop_duplicates().str.cat(sep=", ")).values
        
        file_level_detail["File Path"] = data["File Path"].drop_duplicates().values
        
        # Relevance check -- code here - assuming relevant count = 5
        file_level_detail["Relevant?"] = data[["File Name"]].groupby(["File Name"]).size().\
            reset_index()[0].apply(lambda x: x>5).values
        
        # DRL items satisfied by this 
        
        # writing to "File Level Detail" sheet
        writer = pd.ExcelWriter('demo.xlsx')
        #file_level_detail.to_excel(writer, 'File Level Detail', header = True)  
        with pd.ExcelWriter('demo.xlsx', engine='openpyxl', mode='a') as writer:
            file_level_detail.to_excel(writer, "File Level Detail", index = False)    
        writer.save()

def generatingKeywordDetails(file_path):
    data = pd.read_excel(open(file_path,'rb'), sheet_name="Raw Details")
    if not data.empty:
    #distinct keywords
        keyword_details = data[["Keyword"]].drop_duplicates()
        
        # distinct keywords with the IT category
        keyword_details = data[["Keyword", "IT Category"]]\
            .drop_duplicates().merge(keyword_details, on="Keyword")
            
        #  Number of files with keywords  
        keyword_details = data[["Keyword", "File Name"]].drop_duplicates().groupby(["Keyword"])\
            .size().reset_index().merge(keyword_details, on = "Keyword")
        
        # renaming
        keyword_details = keyword_details.\
            rename(columns={'IT Category': 'Keyword Category', 0: 'Number of files with keywords'})
            
        #Total keyword hits
        keyword_details = data[["File Name","Keyword"]].groupby(["Keyword"])\
            .size().reset_index().merge(keyword_details, on = "Keyword").rename(columns={0:"Total keyword hits"})
    
        #File(s) with maximum keyword hits and File Path(s)
        intermediate2 = data[["Keyword", "File Name", "File Path"]].groupby(["Keyword", "File Name", "File Path"])\
            .size().reset_index().merge(keyword_details, on = "Keyword").rename({0:"Count"}, axis = 1)\
            [["File Name", "Keyword", "File Path", "Count"]]
        
        intermediate3 = intermediate2.groupby('Keyword')['Count'].apply(lambda x : x.eq(x.max()))
        
        intermediate4 = intermediate2.loc[intermediate3].groupby(['Keyword'])['File Name'].agg(', '.join).reset_index()
        
        intermediate5 = intermediate2.loc[intermediate3].groupby(['Keyword'])['File Path'].agg(', '.join).reset_index()
        
        keyword_details = keyword_details.merge(intermediate4, on = "Keyword")
        
        keyword_details = keyword_details.merge(intermediate5, on = "Keyword")
           
        # All file(s) with keyword hits
        keyword_details["All file(s) with keyword hits"] = \
            data[["Keyword", "File Path"]].groupby(["Keyword"])\
            .apply(lambda x : x["File Path"].drop_duplicates().str.cat(sep=", ")).reset_index().\
            merge(keyword_details,on = "Keyword")[[0]].values
            
        # Index Number(s)
        keyword_details["Index Number(s)"] = \
            data[["Keyword", "Index Number"]].groupby("Keyword")\
            .apply(lambda x : x["Index Number"].drop_duplicates().str.cat(sep=", ")).reset_index()\
            .merge(keyword_details, on = "Keyword")[[0]].values   
            
            
        # writing to "File Level Detail" sheet
        writer = pd.ExcelWriter('demo.xlsx')
        #file_level_detail.to_excel(writer, 'File Level Detail', header = True)  
        with pd.ExcelWriter('demo.xlsx', engine='openpyxl', mode='a') as writer:
            keyword_details.to_excel(writer, "Keyword Details", index = False)    
        writer.save()
            
# triggering the application
if __name__ == '__main__':
    wb = xlrd.open_workbook(r"C:\Users\sb512911\Desktop\All\Applications\VDR\Keyword Index for VDR Automation.xlsx")
    sh = wb.sheet_by_index(0)
   
    it_spend = [i for i in sh.col_values(1, start_rowx = 1) if i]
    it_org = [i for i in sh.col_values(2, start_rowx = 1) if i]
    it_apps = [i for i in sh.col_values(3, start_rowx = 1) if i]
    it_infra = [i for i in sh.col_values(4, start_rowx = 1) if i]
    it_security = [i for i in sh.col_values(5, start_rowx = 1) if i]
    it_projects = [i for i in sh.col_values(6, start_rowx = 1) if i]
    target_tech_fitness = [i for i in sh.col_values(7, start_rowx = 1) if i]
    misecellaneous = [i for i in sh.col_values(8, start_rowx = 1) if i]
   
   
    r1 = re.compile('|'.join([r'\b%s\b' % w for w in it_spend]), flags=re.I)
    r2 = re.compile('|'.join([r'\b%s\b' % w for w in it_org]), flags=re.I)
    #r3 = re.compile('|'.join([r'\b%s\b' % w for w in it_apps]), flags=re.I)
    r4 = re.compile('|'.join([r'\b%s\b' % w for w in it_infra]), flags=re.I)
    r5 = re.compile('|'.join([r'\b%s\b' % w for w in it_security]), flags=re.I)
    r6 = re.compile('|'.join([r'\b%s\b' % w for w in it_projects]), flags=re.I)
    r7 = re.compile('|'.join([r'\b%s\b' % w for w in target_tech_fitness]), flags=re.I)
    r8 = re.compile('|'.join([r'\b%s\b' % w for w in misecellaneous]), flags=re.I)
   
    root = r'C:\Users\sb512911\Desktop\All\Applications\VDR\VDR data\RoundTrip Data Room'
    #root = r"C:\Users\sb512911\Desktop\All\Applications\VDR\data\xlsm"
    #csv_handler(file_name)  
   
    fileList = []
    for path, subdirs, files in os.walk(root):
        for name in files:
            a = os.path.join(path, name)
            fileList.append(a)
           
    for i in fileList:
        if i.endswith(".csv"):
            csv_handler(i)   
        if i.endswith((".xlsx", ".xlsm", ".xls")):
            xlsx_handler(i)
        if i.endswith((".xlsb")):
            xlsb_handler(i)
            
    workbook.close()
    
    file_path = r"C:\Users\sb512911\Desktop\All\Applications\VDR\output\demo.xlsx"
    generatingFileLevelDetail(file_path)
    generatingKeywordDetails(file_path)
