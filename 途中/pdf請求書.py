import pyodbc
import pandas as pd
import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

# ✅ 日本語フォント（MSゴシックやIPAフォントなど）を登録
pdfmetrics.registerFont(TTFont('MSGothic', 'C:/Windows/Fonts/msgothic.ttc'))  # Windows用
# pdfmetrics.registerFont(TTFont('IPAexGothic', '/usr/share/fonts/opentype/ipafont-mincho/ipag.ttf'))  # Mac/Linux用

# ユーザーに日付・組勘番号・農家名を入力させる
start_date = input("開始日を YYYY/MM/DD 形式で入力してください: ")
end_date = input("終了日を YYYY/MM/DD 形式で入力してください: ")
group_number = input("組勘番号を入力してください: ")
farmer_name = input("農家名を入力してください: ")

# SQL の日付フォーマットに合わせる
start_date_sql = f"#{start_date}#"
end_date_sql = f"#{end_date}#"

conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\josej\OneDrive\Documents\カルテ.mdb;'
)
conn = pyodbc.connect(conn_str)

# SQLクエリをユーザー入力の日付と組勘番号で動的に作成し、番号順に並べる
sql = f"""
    SELECT 番号, 病名, 金額 
    FROM 平成１５年 
    WHERE 日付 BETWEEN {start_date_sql} AND {end_date_sql} 
    AND 組間番号 = {group_number}
    ORDER BY 番号 ASC
"""

df = pd.read_sql(sql, conn)

# 金額の合計を計算
total_amount = df["金額"].sum()
tax = total_amount * 0.1  # 消費税 10%
grand_total = total_amount + tax

# 今日の日付を取得
today = datetime.date.today().strftime("%Y/%m/%d")

# PDFファイル名
pdf_filename = "invoice.pdf"

# ✅ PDFを作成（フォント適用）
c = canvas.Canvas(pdf_filename, pagesize=A4)
width, height = A4
c.setFont("MSGothic", 12)  # 日本語フォントを指定

# タイトル
c.drawString(50, height - 50, "請求書")

# 農家名と日付
c.drawString(50, height - 80, f"農家名: {farmer_name}")
c.drawString(300, height - 80, f"処理日: {today}")

# ヘッダー
c.drawString(50, height - 110, "番号")
c.drawString(100, height - 110, "病名")
c.drawString(350, height - 110, "金額")

# 横線
c.line(50, height - 115, 500, height - 115)

# データ部分
y_position = height - 140  # 初期Y位置
for index, row in df.iterrows():
    c.drawString(50, y_position, str(row["番号"]))
    c.drawString(100, y_position, row["病名"])  # 日本語フォント適用済み
    c.drawString(350, y_position, f"{row['金額']} 円")
    y_position -= 20  # 行間

# 合計金額・消費税・総合計
y_position -= 30
c.drawString(50, y_position, "合計")
c.drawString(350, y_position, f"{total_amount} 円")

y_position -= 20
c.drawString(50, y_position, "消費税 (10%)")
c.drawString(350, y_position, f"{tax} 円")

y_position -= 20
c.drawString(50, y_position, "総合計")
c.drawString(350, y_position, f"{grand_total} 円")

# PDFを保存
c.save()

# 結果を表示
print(f"PDFファイル '{pdf_filename}' を作成しました。")
print(f"農家名: {farmer_name}")
print(f"処理日: {today}")
print(f"金額の合計: {total_amount} 円")
print(f"消費税 (10%): {tax} 円")
print(f"総合計: {grand_total} 円")
