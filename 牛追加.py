import pyodbc
 
# データベースに接続します
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\tsuboi\OneDrive\Documents\カルテ.mdb;'
    )
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()
#成功


# データベースにデータを追加します
sql = "INSERT INTO テーブル４(生年月日, 農家名, 個体ＩＤ) VALUES(#2021/01/29#, '高橋昌也', 1381439120)"
#sql = "INSERT INTO テ-ブル１(農家名, 個体ＩＤ, 生年月日) VALUES('熊田一夫', 1593512031, #2020/05/30#)"
cursor.execute(sql)
cursor.commit()
 
# データベースの接続を閉じます
cursor.close()
conn.close()
