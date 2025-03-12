import pyodbc
import pandas as pd
import datetime

# ユーザーに日付・組勘番号・農家名を入力させる
start_date = input("開始日を YYYY/MM/DD 形式で入力してください: ")
end_date = input("終了日を YYYY/MM/DD 形式で入力してください: ")
group_number = input("組勘番号を入力してください: ")
farmer_name = input("農家名を入力してください: ")

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

# 消費税（10%）を計算
tax = total_amount * 0.1
grand_total = total_amount + tax

# 今日の日付を取得
today = datetime.date.today().strftime("%Y/%m/%d")

# CSVに保存（合計・消費税・総合計・農家名・日付を追加）
df.loc["合計"] = ["", "合計", total_amount]
df.loc["消費税"] = ["", "消費税 (10%)", tax]
df.loc["総合計"] = ["", "総合計", grand_total]
df.loc["農家名"] = ["", f"農家名: {farmer_name}", ""]
df.loc["処理日"] = ["", f"処理日: {today}", ""]

df.to_csv("employee.csv", encoding="shift_jis", index=False)

# 結果を表示
print("データを employee.csv に出力しました。")
print(f"農家名: {farmer_name}")
print(f"処理日: {today}")
print(f"金額の合計: {total_amount} 円")
print(f"消費税 (10%): {tax} 円")
print(f"総合計: {grand_total} 円")
