#! /usr/bin/python

"""
Excel to SQL v1.0
Python script to convert a Excel to SQL statement
Created by: Miguel Jimenez
Date: 02/11/2014
"""

import sys
import os
import xlrd
import unicodedata

if sys.argv[1] == "help":
    print("Usage:\n\texcelToSQL.py excel_file.xls name_of_excel_page name_of_sql_table")
    sys.exit(1)

if not os.path.exists(sys.argv[1]):
    print("Cannot open "+sys.argv[1])
    sys.exit(1)

filename = sys.argv[1]
extension = filename.split('.')[-1]

if not extension in ('xls'):
    print("The extension of excel_file is incorrect")
    sys.exit(1)

try:
    page = sys.argv[2]
    table = sys.argv[3]
except:
    print("Some arguments are wrong")
    sys.exit(1)

workbook = xlrd.open_workbook(filename)
try:
    worksheet = workbook.sheet_by_name(page)
except:
    print("The page is incorrect")
    sys.exit(1)

sql = "INSERT INTO " + table + "("

j = 0
while j < worksheet.ncols:
    sql += str(worksheet.cell_value(0, j)) + ', '
    j += 1

sql = sql[:len(sql)-2] + ") values"

i = 1
while i < worksheet.nrows:
    sql += '('
    j = 0
    while j < worksheet.ncols:
        try:
            sql += '\'' + str(worksheet.cell_value(i, j)) + '\', '
        except:
             sql += '\'' + unicodedata.normalize('NFKC', worksheet.cell_value(i, j)).encode('utf-8', 'replace') + '\', '
        j += 1
    sql = sql[:len(sql)-2] + '), '
    i += 1
sql = sql[:len(sql)-2] + ';\n'
print sql
sys.exit(0)
