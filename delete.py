import pyodbc

try:
    conn_str =r'Driver={Microsoft Access Driver (*.mdb)};DBQ=E:\GrandSchedule_Coding\GrandSchedule\GrandSchedule.mdb;' #สำหรับ msaccess 32bit + python 32bit
    conn=pyodbc.connect(conn_str)
    crsr=conn.cursor()
    
    dayid=2
    
    crsr.execute('DELETE * FROM tblDay WHERE day_id = ?',dayid)

    print("update :" , crsr.rowcount)
    conn.commit()
    
    
except pyodbc.Error as e:    
    print('error'+ '\n' + str(e))
