from datetime import date, timedelta

# 今日の日付を取得
today = date.today()
print(f"今日の日付: {today}")

# 16日を加算
added_date = today + timedelta(days=16)
print(f"16日後の日付: {added_date}")