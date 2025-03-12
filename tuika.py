import pyodbc
 
# データベースに接続します
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\tsuboi\OneDrive\Documents\カルテ.mdb;'
    )
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()
 
# データベースにデータを追加します
#sql = "INSERT INTO SampleTable(日付, 名前, 支払額) VALUES(#2020/7/19#, '山中一郎', 3900)"
sql = "INSERT INTO テ-ブル４(農家名, 個体ＩＤ, 生年月日) VALUES('熊田一夫', 1593512031, #2020/05/13#)"
cursor.execute(sql)
cursor.commit()
 
# データベースの接続を閉じます
cursor.close()
conn.close()
