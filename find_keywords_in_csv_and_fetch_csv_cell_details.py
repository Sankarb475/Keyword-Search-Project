# -*- coding: utf-8 -*-
"""
Created on Tue Jan 21 14:15:35 2020

@author: sbiswas149
"""
def cell_header(a):
    if a <= 26:
        return chr(a+64)
    count = a//26 + 64
    rest = a%26 + 64
    return chr(count) + chr(rest)

list1 = ['will', 'PwC', 'you']

r1 = re.compile('|'.join([r'\b%s\b' % w for w in list1]), flags=re.I)

with open(r'C:\Users\sbiswas149\Applications\Project\VDR\data.csv') as csvfile:
    readCSV = csv.reader(csvfile)
    count = 0
    for row in readCSV:
        #print(row)
        count = count + 1
        for i in range(len(row)):
            #print(row[i] + " " + str(count) + " " + str(i))
            mm = r1.findall(row[i])
            if mm:
                print(row[i], cell_header(i+1), count)
            

    

