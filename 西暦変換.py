def reiwa_to_seireki(reiwa_year):
    """令和の年を西暦に変換する"""
    if reiwa_year < 1:
        return "令和1年（2019年）以降を入力してください"
    return 2018 + reiwa_year  # 令和1年は2019年

# ユーザー入力
reiwa_year = int(input("令和何年？: "))
print(f"西暦 {reiwa_to_seireki(reiwa_year)} 年")
