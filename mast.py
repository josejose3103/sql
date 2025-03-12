from datetime import datetime, timedelta

# 任意の日付を文字列で入力（例: "2025-02-12"）
date_str = input("日付を YYYY-MM-DD の形式で入力してください: ")

# 文字列を datetime オブジェクトに変換
date_obj = datetime.strptime(date_str, "%Y-%m-%d")

# 14日を加算
future_date = date_obj + timedelta(days=16)

# 結果を表示
print(f"入力された日付: {date_obj.strftime('%Y-%m-%d')}")
print(f"14日後の日付: {future_date.strftime('%Y-%m-%d')}")
print("Python スクリプトが正常に実行されました！")
input("Press Enter to exit...")