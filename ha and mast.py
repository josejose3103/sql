import datetime

# 現在時刻
dt1 = datetime.datetime.now()
print(dt1)
dt2 = dt1 + datetime.timedelta(days=21)
dt3 = dt1 + datetime.timedelta(days=14)
print(dt2)
print(dt3)
