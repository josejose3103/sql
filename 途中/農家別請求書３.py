import pyodbc
import pandas as pd

# ユーザーに日付と組勘番号を入力させる
start_date = input("開始日を YYYY/MM/DD 形式で入力してください: ")
end_date = input("終了日を YYYY/MM/DD 形式で入力してください: ")
group_number = input("組勘番号を入力してください: ")

# SQL の日付フォーマットに合わせる
start_date = f"#{start_date}#"
end_date = f"#{end_date}#"

conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\josej\OneDrive\Documents\カルテ.mdb;'
)
conn = pyodbc.connect(conn_str)

# SQLクエリをユーザー入力の日付と組勘番号で動的に作成し、番号順に並べる
sql = f"""
    SELECT 番号, 病名, 金額 
    FROM 平成１５年 
    WHERE 日付 BETWEEN {start_date} AND {end_date} 
    AND 組間番号 = {group_number}
    ORDER BY 番号 ASC
"""

df = pd.read_sql(sql, conn)

# CSVに保存
df.to_csv("employee.csv", encoding="shift_jis", index=False)

print("データを employee.csv に出力しました。")
