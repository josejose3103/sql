import csv
import MySQLdb

# データベース接続(文字化けしましたので　utf8指定)
conn = mysql.connector.connect(
    host="localhost",
    user="root",
    password="3103@Kazu",
    database="media",
    charset="utf8"
)
cursor = conn.cursor()

# ユーザー入力の取得
start_date = input("開始日 (YYYY-MM-DD): ")
end_date = input("終了日 (YYYY-MM-DD): ")

# SQLクエリ
query = """
    SELECT noukamei, SUM(kingaku)
    FROM article3
    WHERE day BETWEEN %s AND %s
    GROUP BY noukamei;
"""
cursor.execute(query, (start_date, end_date))
results = cursor.fetchall()

# CSVファイルの作成
csv_filename = "output.csv"

# 書き込むデータを準備する
with open(csv_filename, 'w') as f:
    writer = csv.writer(f)
    writer.writerow(['農家名', '合計金額'])
    
    for row in results:
        noukamei, kingaku = row
        writer.writerow([str(noukamei), str(kingaku)])

# データベースを閉じる
conn.close()

print(f"CSVファイル '{csv_filename}' を作成しました。")
