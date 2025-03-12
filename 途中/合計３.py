import pyodbc
import pandas as pd

# ユーザーから日付範囲を入力
start_date = input("開始日を入力 (YYYY/MM/DD): ")
end_date = input("終了日を入力 (YYYY/MM/DD): ")

# SQLクエリの作成（農家名ごとの金額合計）
sql = f"""
    SELECT 農家名, SUM(金額) AS 金額の合計
    FROM 平成１５年
    WHERE 日付 BETWEEN #{start_date}# AND #{end_date}#
    GROUP BY 農家名
    ORDER BY 金額の合計 DESC
"""

# データベース接続
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\josej\OneDrive\Documents\カルテ.mdb;'
)
conn = pyodbc.connect(conn_str)

df = pd.read_sql(sql, conn)
print(df)

# 必要ならCSVに保存
# df.to_csv("syukei.csv", encoding="shift_jis", index=False)
