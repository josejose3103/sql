from dateutil.relativedelta import relativedelta

date = datetime.date(2021, 2, 18)
print(date) # 2018-05-01

# 1ヶ月後を求める
print(date + relativedelta(months=1)) # 2018-06-01

# 1ヶ月前を求める
print(date + relativedelta(months=-1)) # 2018-04-01

# 1日後を求める
print(date + relativedelta(days=21)) # 2018-05-02
