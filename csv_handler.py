# -*- coding: utf-8 -*-
"""
Created on Mon Jan 20 12:36:08 2020

@author: sbiswas149
"""

import csv
import xlrd
import re
import os
import xlsxwriter

outputCol = 0
outputRow = 1

os.chdir(r"C:\Users\sb512911\Desktop\All\Applications\VDR\output")
workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet("Raw Details")
bold = workbook.add_format({'bold': True})
worksheet.write('A1','File Name', bold)
worksheet.write('B1','Index Number', bold)
worksheet.write('C1','Folder Name', bold)
worksheet.write('D1','File Path', bold)
worksheet.write('E1','File Type', bold)
worksheet.write('F1','Total Number of Page(s)/Slide(s)', bold)
worksheet.write('G1','Keyword', bold)
worksheet.write('H1','IT Category', bold)
worksheet.write('I1','Page/Slide Number', bold)

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
        for row in readCSV:
            count = count + 1
            for i in range(len(row)):
                #print(row[i] + " " + str(count) + " " + str(i))
                mm = r1.findall(row[i])
                if mm:
                    list_details = folderName(path)
                    cell_details = cell_header(i+1)+str(count)
                    outputWriter(list_details[0], list_details[1], list_details[2], path, list_details[3], 1,
                                 row[i], "IT Spend", cell_details)

def outputWriter(file_name, index_number, folder_name, file_path, file_type, pages, keyword, category, cell):
    global outputRow
    global outputCol
    worksheet.write(outputRow,0,file_name)
    worksheet.write(outputRow,1,index_number)
    worksheet.write(outputRow,2,folder_name)
    worksheet.write(outputRow,3,file_path)
    worksheet.write(outputRow,4,file_type)
    worksheet.write(outputRow,5,pages)
    worksheet.write(outputRow,6,keyword)
    worksheet.write(outputRow,7,category)
    worksheet.write(outputRow,8,cell)
    outputRow = outputRow + 1
   
def folderName(path):
    #C:\Users\sb512911\Desktop\All\Screenshots\1.2.1 ILS Structure Chart.csv
    file_name = path.split("\\")[-1]
    index_number = file_name.split(" ")[0]
    folder_name = path.split(file_name)[0]
    file_type = path.split(".")[-1]   
    return [file_name, index_number, folder_name, file_type]
   
# triggering the ETL application
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
   
    root = r'C:\Users\sb512911\Desktop\All\Applications\VDR\data'
    #csv_handler(file_name)  
   
    fileList = []
    for path, subdirs, files in os.walk(root):
        for name in files:
            a = os.path.join(path, name)
            fileList.append(a)
           
    for i in fileList:
        if i.endswith(".csv"):
            csv_handler(i)
           
           
    workbook.close()
