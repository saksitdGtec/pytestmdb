import pyodbc

conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb)};DBQ=E:\GrandSchedule_Coding\GrandSchedule\GrandSchedule.mdb;') #สำหรับ msaccess 32bit + python 32bit
cursor = conn.cursor()
cursor.execute(r'select * from tblDay where day_id = 2')

for row in cursor.fetchval():
    print (str(row))