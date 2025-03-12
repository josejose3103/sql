from datetime import datetime, timedelta

# 今日の日付
today = datetime.today()

# 16日後の日付
future_date = today + timedelta(days=16)

# 結果を表示
print(future_date.strftime("%Y-%m-%d"))
