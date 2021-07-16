import pyodbc

try:
    #conn_str =r'Driver={Microsoft Access Driver (*.mdb)};DBQ=E:\GrandSchedule_Coding\GrandSchedule\GrandSchedule.mdb;' #สำหรับ msaccess 32bit + python 32bit
    conn_str =r'Driver={Microsoft Access Driver (*.mdb)};DBQ=C:\Users\Administrator\Desktop\22att2000\22att2000.mdb;' #สำหรับ msaccess 32bit + python 32bit
    conn= pyodbc.connect(conn_str)
    crsr= conn.cursor()
    userid=106
    crsr.execute('select * from abk_userinfo where userid = ?',userid)
    rows = crsr.fetchall()
    for row in rows:
        print(row)

    result="ok"
    print(f"connect to dbs {result}")

except pyodbc.Error as e:
    myerr="hi error"
    print(f"error {myerr} : ",e)