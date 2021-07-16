import pyodbc

conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb)};DBQ=E:\GrandSchedule_Coding\GrandSchedule\GrandSchedule.mdb;') #สำหรับ msaccess 32bit + python 32bit
cursor = conn.cursor()
cursor.execute('select * from tblDay')

for row in cursor.fetchall():
    print (row)