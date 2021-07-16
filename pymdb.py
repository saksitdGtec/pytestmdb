
# https://github.com/mkleehammer/pyodbc/wiki/Connecting-to-Microsoft-Access


import pyodbc


conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb)};DBQ=E:\GrandSchedule_Coding\GrandSchedule\GrandSchedule.mdb;') #สำหรับ msaccess 32bit + python 32bit

#สำหรับ py64bit + msaccess 64bit
# conn_str = (
#     r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
#     r'DBQ=C:\path\to\mydb.accdb;'
#     )

cursor = conn.cursor()
cursor.execute(r'select * from tblDay where day_id = 2')

for row in cursor.fetchval():
    print (str(row))