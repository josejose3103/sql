import pyodbc

# データベースに接続
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\josej\OneDrive\Documents\カルテ.mdb;'
)
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# ユーザーから入力を取得
フィールド1 = input("日付を入力してください (YYYY-MM-DD): ")
フィールド2 = input("診療損害防止等の内容を入力してください: ")
フィールド3 = input("金額を入力してください: ")
フィールド4 = input("その他の情報を入力してください: ")
組勘 = input("組勘を入力してください: ")
# SQLクエリをパラメータ化して実行
sql = "INSERT INTO 保険外　薬品 (フィールド1, フィールド2, フィールド3, フィールド4, 組勘) VALUES (?, ?, ?, ?, ?)"
cursor.execute(sql, (フィールド1, フィールド2, フィールド3, フィールド4, 組勘))

# コミットして保存
cursor.commit()

# データベースの接続を閉じる
cursor.close()
conn.close()

print("データが正常に追加されました。")
