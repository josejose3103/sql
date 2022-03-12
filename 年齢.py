# PYTHON_AGE_01-1

from datetime import date
from dateutil.relativedelta import relativedelta
# PYTHON_AGE_01-2

# 生年月日
d0 = date(2013, 1, 25)

# 現在の日付
d1 = date.today()

# 経過時間
dy = relativedelta(d1, d0)

print(dy)
