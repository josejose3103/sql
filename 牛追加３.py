import pyodbc

# データベースに接続
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\josej\OneDrive\Documents\カルテ.mdb;'
)
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# SQL文
sql = "INSERT INTO テーブル４ (生年月日, 農家名, 個体ＩＤ) VALUES (?, ?, ?)"

# ユーザーからデータを入力
while True:
    生年月日 = input("生年月日 (YYYY/MM/DD): ")
    農家名 = input("農家名: ")
    個体ＩＤ = input("個体ID: ")

    # データ挿入
    cursor.execute(sql, (生年月日, 農家名, int(個体ＩＤ)))
    conn.commit()

    # 続けるか確認
    続ける = input("続けますか？ (yes/no): ").lower()
    if 続ける != 'yes':
        break

# データベースの接続を閉じる
cursor.close()
conn.close()
