
import pyodbc
import pandas as pd
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\tsuboi\OneDrive\Documents\カルテ.mdb;'
    )
conn = pyodbc.connect(conn_str)

sql = "SELECT sum(金額) FROM 平成１５年 WHERE 日付 BETWEEN #2022/01/10# AND #2022/01/25#"
#sql = 'SELECT 農家名,番号,病名,金額 FROM 平成１５年 WHERE 日付 = #2020/12/12#'
df = pd.read_sql(sql, conn)
print(df)
#df.to_csv("syukei.csv", encoding="shift_jis")




 
