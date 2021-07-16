import pyodbc
import datetime

try:
    
    conn_str =r'Driver={Microsoft Access Driver (*.mdb)};DBQ=E:\GrandSchedule_Coding\GrandSchedule\GrandSchedule.mdb;' #สำหรับ msaccess 32bit + python 32bit
    conn=pyodbc.connect(conn_str)
    crsr=conn.cursor()
    dayname="daytest3"
    dayid=2
    crsr.execute('UPDATE tblDay SET day_name = ? WHERE day_id = ?',(dayname,dayid))
    conn.commit()
    
    result="ok"
    print(f'update {result}')
    
except:
    print('error')