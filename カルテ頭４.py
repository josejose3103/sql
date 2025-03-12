import pyodbc
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4

# ユーザー入力
custom_text = input("住所の3cm下に表示する文字を入力してください: ")

# ユーザー入力で検索する個体ID
individual_id = input("検索する個体IDを入力してください: ")

# データベース接続情報
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\josej\OneDrive\Documents\カルテ.mdb;'
)

conn = pyodbc.connect(conn_str)

# SQLクエリの作成
sql = "SELECT 農家名, 個体ＩＤ, 生年月日 FROM テーブル４ WHERE 個体ＩＤ=?"
df = pd.read_sql(sql, conn, params=[individual_id])

# 生年月日を "YYYY/MM/DD" 形式に変換
df["生年月日"] = pd.to_datetime(df["生年月日"]).dt.strftime("%Y/%m/%d")

# PDF作成
pdf_filename = f"個体_{individual_id}.pdf"
c = canvas.Canvas(pdf_filename, pagesize=A4)

# 日本語フォントを登録
pdfmetrics.registerFont(TTFont('Gothic', 'C:/Windows/Fonts/msgothic.ttc'))  # Windows

c.setFont("Gothic", 12)  # フォントサイズ設定

# A4 サイズ
page_width, page_height = A4

y_position = 807  # 約2cm（57pt）上に調整

# **1回だけ住所・施設名・名前を出力**
c.drawString(50, y_position, "弟子屈町朝日２丁目4-30")
c.drawString(50, y_position - 20, "坪井家畜診療所")
c.drawString(50, y_position - 40, "坪井聡之")

# **テキストの位置調整**
text_y_position = y_position - 40 - 85  # 住所から3cm下
c.drawString(85, text_y_position - 28.35, custom_text)  # 左から3cm・1cm下

# **1回だけ農家名・個体ID・生年月日を出力**
farmer_name = df.iloc[0]["農家名"]
individual_id = str(df.iloc[0]["個体ＩＤ"])
birth_date = df.iloc[0]["生年月日"]

# 農家名を **右揃え**
c.setFont("Gothic", 16)  # 農家名は少し大きめ
c.drawRightString(page_width - 50, y_position, farmer_name)

# 個体IDと生年月日を **右揃え**
c.setFont("Gothic", 14)
c.drawRightString(page_width - 50, y_position - 30, f"{individual_id}  {birth_date}")

c.save()
conn.close()

print(f"PDFファイル '{pdf_filename}' が作成されました。")
