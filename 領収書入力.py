import pyodbc
import datetime

# データベースに接続
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\josej\OneDrive\Documents\カルテ.mdb;'
)
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# ユーザー入力を取得
組間番号 = int(input("組間番号を入力してください: "))  # 整数型に変換
農家名 = input("農家名を入力してください: ")  # 文字列
番号 = int(input("番号を入力してください: "))  # 整数型に変換
病名 = input("病名を入力してください: ")  # 文字列
金額 = float(input("金額を入力してください: "))  # 金額なので float に変換
日付 = input("日付を YYYY/MM/DD 形式で入力してください: ")

# 日付を datetime.date 型に変換
日付 = datetime.datetime.strptime(日付, "%Y/%m/%d").date()

# SQLクエリを作成（パラメータ化クエリを使用）
sql = "INSERT INTO 平成15年 (組間番号, 農家名, 番号, 病名, 金額, 日付) VALUES (?, ?, ?, ?, ?, ?)"

# クエリを実行
cursor.execute(sql, (組間番号, 農家名, 番号, 病名, 金額, 日付))
cursor.commit()

# データベースの接続を閉じる
cursor.close()
conn.close()

print("データが正常に追加されました。")
