#nothing yet
import xlrd
import os
import re
import nltk
from nltk.tag import pos_tag

def csv_from_excel():
    wb = xlrd.open_workbook('VDR_Index.xlsx')
    sh = wb.sheet_by_name('Sheet1')
    dict1 = {}
    file_name_list = sh.col_values(5, start_rowx=3)
    keyword_list = sh.col_values(9, start_rowx=3)
    #print(type(keyword_list[0]))
    for row in range(1,len(file_name_list)):
        list1 = keyword_list[row][10:-1].split(",")
        for a in range(len(list1)):
            list1[a] = list1[a].replace("'","").strip()
        dict1[file_name_list[row]] = list1
    #print(dict1)
    return dict1


# runs the csv_from_excel function:
workbook = xlrd.open_workbook("DRL.xlsx")
sh = workbook.sheet_by_index(0)
satisfied_output = {}
output = csv_from_excel()
print(output)
for rownum in range(sh.nrows):
    propernouns = []
    sentence = sh.row_values(rownum)[6]
    drl_number = sh.row_values(rownum)[1]
    #print(drl_number)
    #print(sentence)
    #tagged_sent = pos_tag(sentence.split())
    #propernouns = list(set([removeChar(word).strip() for word,pos in tagged_sent if pos == 'NN']))
    #print(rownum)
    #print(propernouns)
    #output = csv_from_excel()
    for keys,values in output.items():
        count = 0
        #for i in propernouns:
            #i = removeChar(i)
        for m in values:
            if m.lower() in sentence.lower():
                count = count + 1
        if count >= 1:
            #print(count, drl_number, keys)
            if drl_number not in satisfied_output:
                satisfied_output[int(drl_number)] =  keys + ''
            else:
                satisfied_output[int(drl_number)] = satisfied_output[int(drl_number)] + keys + " "

print(satisfied_output)
