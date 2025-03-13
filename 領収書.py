import pyodbc
 
# データベースに接続します
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\josej\OneDrive\Documents\カルテ.mdb;'
    )
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()
#成功 
# データベースにデータを追加します
#sql = "INSERT INTO 平成15年(組間番号,農家名, 番号, 病名, 金額, 日付) VALUES(13, '船坂明美', 141620610, '卵巣静止', 6124, #2020/12/16#)"
sql = "INSERT INTO 保険外　薬品(フィールド1, フィールド2, フィールド3, フィールド４) VALUES('2025/03/08', '乾乳用セファロニュームＤ＊５', 7500, '坂上孝行')"

#sql = "INSERT INTO テ-ブル１(農家名, 個体ＩＤ, 生年月日) VALUES('熊田一夫', 1593512031, #2020/05/30#)"
cursor.execute(sql)
cursor.commit()
 
# データベースの接続を閉じます
cursor.close()
conn.close()
