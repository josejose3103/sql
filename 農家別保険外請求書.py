import pyodbc
import pandas as pd
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\josej\OneDrive\Documents\カルテ.mdb;'
    )
conn = pyodbc.connect(conn_str)
#sql = 'SELECT * FROM 平成１５年'
sql = "SELECT 番号,病名,金額 FROM 平成１５年 WHERE 日付 BETWEEN #2020/11/26# AND #2020/12/10# AND 農家名 = '坂上孝行'"
#sql = "SELECT 番号,病名,金額 FROM 平成１５年 WHERE 日付 BETWEEN #2022/01/10# AND #2022/01/25# AND 組間番号 = 12"
#sql = "SELECT 農家名, SUM(金額) FROM 平成１５年 WHERE 日付 BETWEEN #2024/01/10# AND #2024/01/25# GROUP BY 農家名, 組間番号 ORDER BY 組間番号"
# アクセスｓｑｌそのままコピー
#sql = "SELECT 農家別.番号, 農家別.病名, 農家別.金額, 農家別.日付 FROM 農家別 WHERE (((農家別.日付)>#1/10/2024# And (農家別.日付)<=#1/25/2024#))) ORDER BY 農家別.日付"
#sql = "SELECT フィールド1, フィールド４,フィールド2,フィールド3,組勘 FROM 保険外　薬品 WHERE (((フィールド1)>#01/01/2025# And (フィールド1)<=#2/28/2025#)) "
sql = "SELECT フィールド1, フィールド４,フィールド2,フィールド3,組勘 FROM 保険外　薬品 WHERE フィールド1  BETWEEN #2025/01/01# AND #2025/02/28# AND フィールド４ = '坂上孝行'"
#sql = "SELECT フィールド1, フィールド４,フィールド2,フィールド3,組勘 FROM 保険外　薬品 WHERE フィールド1 BETWEEN #2022/01/10# AND #2022/01/25# AND 組間 = 12"


#sql = "SELECT Sum(保険外　薬品.[フィールド3]) AS フィールド3の合計 FROM 保険外, 保険外　薬品 WHERE (((保険外　薬品.[フィールド1]) Between #12/1/2023# And #1/31/2024#));"

df = pd.read_sql(sql, conn)
print(df)

#df.to_csv("employee.csv", encoding="shift_jis")
