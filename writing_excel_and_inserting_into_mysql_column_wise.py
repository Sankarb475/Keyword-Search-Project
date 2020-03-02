# -*- coding: utf-8 -*-
"""
Created on Fri Feb 28 15:36:57 2020

@author: sbiswas149
"""

import pandas as pd
import mysql.connector

mydb = mysql.connector.connect(
      host="localhost",
      user="root",
      passwd="1097",
      database="PyMy"
)

def isNaN(num):
    return num != num

def xlsx_handler(path):
    sheets_dict = pd.read_excel(path, sheet_name="Sheet1")
    dictionary = ["IT Spend", "IT Organization / Roles","IT Applications","IT Infrastructure","IT Security and Controls","IT Projects",
                  "Target Technology Fitness Scorecard","Misecellaneous"]
    
    for index, row in sheets_dict.iterrows():
        for i in range(1,len(row)):
            if not isNaN(row[i]):
                mysqlWriter(row[i], dictionary[i-1])
                
                
def mysqlWriter(keyword, category):
    mycursor = mydb.cursor()
    
    sql = "INSERT INTO data_dictionary (Keyword, Category) VALUES (%s, %s)"
    val = (keyword, category)
    mycursor.execute(sql, val)
    
    mydb.commit()
    
               
path = r"C:\Users\sbiswas149\Applications\data\VDR\keyword_index.xlsx"
xlsx_handler(path)
