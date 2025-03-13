import pyodbc
import pandas as pd
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\tsuboi\OneDrive\Documents\カルテ.mdb;'
    )
conn = pyodbc.connect(conn_str)
sql = 'SELECT * FROM 平成１５年'
#sql = "SELECT 番号,病名,金額 FROM 平成１５年 WHERE 日付 BETWEEN #2020/11/26# AND #2020/12/10# AND 農家名 = '坂上孝行'"
sql = "SELECT 番号,病名,金額 FROM 平成１５年 WHERE 日付 BETWEEN #2020/11/26# AND #2020/12/10# AND 組間番号 = 12"
#sql = 'SELECT 農家名,番号,病名,金額 FROM 平成１５年 WHERE 日付 = #2020/12/12#'
df = pd.read_sql(sql, conn)
print(df)
