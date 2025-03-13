import pyodbc
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import os

# 日本語フォントの設定
font_path = "C:/Windows/Fonts/msgothic.ttc"  # Windowsの場合（MSゴシック）
pdfmetrics.registerFont(TTFont("MSGothic", font_path))

# ユーザーから日付範囲を入力
start_date = input("開始日を入力 (YYYY-MM-DD): ")
end_date = input("終了日を入力 (YYYY-MM-DD): ")

# データベース接続
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\josej\OneDrive\Documents\カルテ.mdb;'
)
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# SQLクエリの作成
sql = """
    SELECT 農家名, SUM(金額) AS 金額の合計
    FROM 平成１５年
    WHERE 日付 BETWEEN ? AND ?
    GROUP BY 農家名
    ORDER BY SUM(金額) DESC
"""

# クエリを実行
cursor.execute(sql, [start_date, end_date])
rows = cursor.fetchall()

# データをデータフレームに変換
df = pd.DataFrame.from_records(rows, columns=[desc[0] for desc in cursor.description])

# 金額を円表示にフォーマット
df["金額の合計"] = df["金額の合計"].apply(lambda x: f"¥{x:,.0f}")

# PDFファイルの作成
pdf_filename = "農家別金額集計表.pdf"
pdf = canvas.Canvas(pdf_filename, pagesize=A4)
pdf.setTitle("農家別金額集計表")

# 日本語フォントを適用
pdf.setFont("MSGothic", 16)

# タイトルを設定
pdf.drawString(200, 800, "農家別金額集計表")

# 表のデータを作成
table_data = [["農家名", "金額の合計"]] + df.values.tolist()

# Tableオブジェクトを作成
table = Table(table_data, colWidths=[250, 150])
table.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ("FONTNAME", (0, 0), (-1, -1), "MSGothic"),  # 日本語フォント適用
    ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
    ("BACKGROUND", (0, 1), (-1, -1), colors.beige),
    ("GRID", (0, 0), (-1, -1), 1, colors.black),
]))

# 表を配置
table.wrapOn(pdf, 50, 500)
table.drawOn(pdf, 50, 700 - len(df) * 20)

# PDFを保存
pdf.save()

# 接続を閉じる
cursor.close()
conn.close()

print(f"PDFファイル '{pdf_filename}' を作成しました！")
