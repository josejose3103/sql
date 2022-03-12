import pandas as pd
from fbprophet import Prophet
import pandas_datareader as pdr
import matplotlib.pyplot as plt

df = pdr.get_data_yahoo("AAPL", "2019-11-01", "2021-11-29")
df = pd.DataFrame({"ds": df.index.values, "y": df["Close"].values})
m = Prophet()
m.fit(df)
# 25日先まで予測
future = m.make_future_dataframe(periods=25, freq="d")
# 土日除外
future = future[future["ds"].dt.weekday < 5]
forecast = m.predict(future)
m.plot(forecast)
plt.savefig("pred.png")
