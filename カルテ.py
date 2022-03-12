import pyodbc
import pandas as pd
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\tsuboi\OneDrive\Documents\カルテ.mdb;'
    )
conn = pyodbc.connect(conn_str)
#cursor = conn.cursor()
sql = "SELECT 農家名,個体ＩＤ,生年月日 FROM テーブル４ WHERE 個体ＩＤ='1397219167'"
#sql = 'SELECT * FROM テ-ブル４ WHERE 個体ＩＤ=1593512031'
#

#cursor.commit()
#cursor.execute(sql)

df = pd.read_sql(sql, conn)
print(df)

