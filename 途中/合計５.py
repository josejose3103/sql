import pyodbc
import pandas as pd

# ユーザーから日付範囲を入力（Accessでは YYYY-MM-DD 形式が確実）
start_date = input("開始日を入力 (YYYY-MM-DD): ")
end_date = input("終了日を入力 (YYYY-MM-DD): ")

# データベース接続
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\josej\OneDrive\Documents\カルテ.mdb;'
)
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# SQLクエリの作成（プレースホルダーを使用）
sql = """
    SELECT 農家名, SUM(金額) AS 金額の合計
    FROM 平成１５年
    WHERE 日付 BETWEEN ? AND ?
    GROUP BY 農家名
    ORDER BY SUM(金額) DESC
"""

# クエリを実行
cursor.execute(sql, [start_date, end_date])
rows = cursor.fetchall()

# 結果をデータフレームに変換
df = pd.DataFrame.from_records(rows, columns=[desc[0] for desc in cursor.description])

# 結果を表示
print(df)

# 必要ならCSVに保存
# df.to_csv("syukei.csv", encoding="shift_jis", index=False)

# 接続を閉じる
cursor.close()
conn.close()
