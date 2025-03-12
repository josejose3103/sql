import pyodbc
import pandas as pd

# データベースに接続
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\josej\OneDrive\Documents\カルテ.mdb;'
)
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# ユーザーに個体IDの一部を入力させる
部分個体ＩＤ = input("検索する個体IDの一部（4桁）を入力してください: ")

# SQLクエリ（個体ID部分一致検索）
sql = "SELECT 個体ＩＤ, 農家名, 生年月日 FROM テーブル４ WHERE 個体ＩＤ LIKE ?"
cursor.execute(sql, (f"%{部分個体ＩＤ}%",))

# 結果を取得
候補リスト = cursor.fetchall()

# 候補を表示
if not 候補リスト:
    print("該当するデータが見つかりませんでした。")
else:
    print("\n該当する個体データ:")
    print("番号 | 個体ID       | 農家名      | 生年月日")
    print("-" * 40)
    for idx, (個体ＩＤ, 農家名, 生年月日) in enumerate(候補リスト):
        print(f"{idx + 1}  | {個体ＩＤ} | {農家名} | {生年月日}")

    # ユーザーに選択させる
    選択番号 = int(input("\n選択する番号を入力してください（キャンセルは0）: ")) - 1

    if 0 <= 選択番号 < len(候補リスト):
        確定個体ＩＤ, 確定農家名, 確定生年月日 = 候補リスト[選択番号]
        print(f"\n選択した個体ID: {確定個体ＩＤ}")
        print(f"農家名: {確定農家名}")
        print(f"生年月日: {確定生年月日}")
    else:
        print("\nキャンセルしました。")

# データベースの接続を閉じる
cursor.close()
conn.close()
