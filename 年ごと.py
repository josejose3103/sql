import pyodbc
import pandas as pd

conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\josej\OneDrive\Documents\カルテ.mdb;'
    )
conn = pyodbc.connect(conn_str)

#sql = "SELECT 農家名, sum(金額) FROM 平成１５年 WHERE 日付 BETWEEN #2020/12/11# AND #2020/12/20# group by 農家名"
#sql = "SELECT 農家名, sum(金額) FROM 平成１５年 WHERE 日付 BETWEEN #2020/12/21# AND #2021/01/10# group by 農家名"
#sql = "SELECT DATEPART(YEAR, 日付), sum(金額) FROM 平成１５年 GROUP BY DATEPART(YEAR, 日付)"
     


#sql = "SELECT sum(金額) FROM 平成１５年 WHERE 日付 BETWEEN #2020/12/21# AND #2021/01/10#"
#sql = 'SELECT 農家名,番号,病名,金額 FROM 平成１５年 WHERE 日付 = #2020/12/12#'
#$sql = "SELECT DATE_FORMAT(regist_time, '%Y-%m') as regist_time, sum(金額) as 金額 FROM 平成１５年 GROUP BY DATE_FORMAT(regist_time, '%Y%m')"
sql = "select year(日付),avg(金額), sum(金額), count(金額) from 平成１５年 group by year(日付) "
#sql = "select year(月日),avg(金額), sum(金額), count(金額) from 保険外　薬品 group by year(月日) "
df = pd.read_sql(sql, conn)
print(df)
#df.to_csv("syukei.csv", encoding="shift_jis")
