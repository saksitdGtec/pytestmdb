import pyodbc
import datetime

try:

    conn_str= conn_str =r'Driver={Microsoft Access Driver (*.mdb)};DBQ=E:\GrandSchedule_Coding\GrandSchedule\GrandSchedule.mdb;' #สำหรับ msaccess 32bit + python 32bit
    conn = pyodbc.connect(conn_str)
    crsr= conn.cursor()
    dt=datetime.datetime.today() # Or datetime.datetime.now()
    
    # สำหรับ เพิ่ม รายรายการ ใช้ list of tuples รันด้วย conn.executemany
    # newday= (
    #             (dt ,'วันเทส2'),
    #             (dt ,'วันเทส3')   
    #     )
    newday= (
                (dt ,'วันเทส2')
        )
    

    

    crsr.execute('INSERT INTO tblDay (day_date,day_name) VALUES (?,?)',newday)
    # crsr.executemany('INSERT INTO tblDay (day_date,day_name) VALUES (?,?)',newday)
    conn.commit()
    print('data inserted')

except pyodbc.Error as e:
    print("error")    