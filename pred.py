import numpy as np
from sklearn.preprocessing import StandardScaler
from sklearn.model_selection import train_test_split
from keras import Sequential
from keras.layers import Dense, GRU
import pandas_datareader as pdr
from sklearn.metrics import accuracy_score

df = pdr.get_data_yahoo("AAPL", "2010-11-01", "2020-11-01")
df["Diff"] = df.Close.diff()
df["SMA_2"] = df.Close.rolling(2).mean()
df["Force_Index"] = df.Close * df.Volume
df["y"] = df["Diff"].apply(lambda x: 1 if x > 0 else 0).shift(-1)
df = df.drop(
   ["Open", "High", "Low", "Close", "Volume", "Diff", "Adj Close"],
   axis=1,
).dropna()
# print(df)
X = StandardScaler().fit_transform(df.drop(["y"], axis=1))
y = df["y"].values
X_train, X_test, y_train, y_test = train_test_split(
   X,
   y,
   test_size=0.2,
   shuffle=False,
)
model = Sequential()
model.add(GRU(2, input_shape=(X_train.shape[1], 1)))
model.add(Dense(1, activation="sigmoid"))
model.compile(optimizer="adam", loss="binary_crossentropy", metrics=["acc"])
model.fit(X_train[:, :, np.newaxis], y_train, epochs=100)
y_pred = model.predict(X_test[:, :, np.newaxis])
print(accuracy_score(y_test, y_pred > 0.5))