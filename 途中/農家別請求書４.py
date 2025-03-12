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

# 金額の合計を計算
total_amount = df["金額"].sum()

# CSVに保存（合計行を追加）
df.loc["合計"] = ["", "合計", total_amount]  # "番号" は空欄, "病名" に "合計" を入れる

df.to_csv("employee.csv", encoding="shift_jis", index=False)

print("データを employee.csv に出力しました。")
print(f"金額の合計: {total_amount} 円")
