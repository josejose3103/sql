import pyodbc
import pandas as pd
import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

# 📅 今日の日付を取得（YYYYMMDD 形式）
today_str = datetime.datetime.today().strftime("%Y%m%d")

# 📄 日本語フォントの設定（メイリオを使用）
font_path = "C:/Windows/Fonts/meiryo.ttc"
pdfmetrics.registerFont(TTFont("Meiryo", font_path))

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

# SQLクエリ（船坂昭美を除外）
sql = """
    SELECT 農家名, SUM(金額) AS 金額の合計
    FROM 平成１５年
    WHERE 日付 BETWEEN ? AND ? AND 農家名 <> '船坂昭美'
    GROUP BY 農家名
    ORDER BY SUM(金額) DESC
"""

# クエリを実行
cursor.execute(sql, [start_date, end_date])
rows = cursor.fetchall()

# データをデータフレームに変換
df = pd.DataFrame.from_records(rows, columns=[desc[0] for desc in cursor.description])

# 総合計を計算
total_amount = df["金額の合計"].sum()

# 金額を円表示にフォーマット
df["金額の合計"] = df["金額の合計"].apply(lambda x: f"¥{x:,.0f}")

# 総合計をデータフレームに追加
df.loc["総合計"] = ["総合計", f"¥{total_amount:,.0f}"]

# 📄 PDFファイルの作成（出力日付き）
pdf_filename = f"農家別金額集計表_{today_str}_除外済.pdf"
pdf = canvas.Canvas(pdf_filename, pagesize=A4)
pdf.setTitle("農家別金額集計表")

# 日本語フォントを適用
pdf.setFont("Meiryo", 12)

# 左上に「保険診療集計表」
pdf.drawString(50, 820, "保険診療集計表")

# 右上に「坪井家畜診療所」
pdf.drawString(400, 820, "坪井家畜診療所")

# タイトルの下にお知らせ文
message = f"{start_date} から {end_date} までの請求分保険診療費は以下の通りですのでお知らせいたします。（船坂昭美を除外）"
pdf.drawString(50, 800, message)

# タイトルの下に少しスペースを開ける
y_position = 770

# 表のデータを作成
table_data = [["農家名", "金額の合計"]] + df.values.tolist()

# Tableオブジェクトを作成
table = Table(table_data, colWidths=[250, 150])
table.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.grey),  # ヘッダー背景色
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),  # ヘッダー文字色
    ("ALIGN", (0, 0), (-1, -1), "CENTER"),  # すべて中央揃え
    ("FONTNAME", (0, 0), (-1, -1), "Meiryo"),  # メイリオフォント適用
    ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
    ("BACKGROUND", (0, 1), (-1, -2), colors.beige),  # 通常データの背景色
    ("GRID", (0, 0), (-1, -1), 1, colors.black),  # グリッド線

    # 総合計の行のデザイン
    ("BACKGROUND", (0, -1), (-1, -1), colors.lightgrey),  # 総合計の背景色
    ("TEXTCOLOR", (0, -1), (-1, -1), colors.black),  # 総合計の文字色
    ("FONTNAME", (0, -1), (-1, -1), "Meiryo"),  # メイリオフォント適用
    ("ALIGN", (0, -1), (-1, -1), "CENTER"),  # 総合計の中央揃え
]))

# 表を配置
table.wrapOn(pdf, 50, 500)
table.drawOn(pdf, 50, y_position - len(df) * 20)

# PDFを保存
pdf.save()

# 接続を閉じる
cursor.close()
conn.close()

print(f"PDFファイル '{pdf_filename}' を作成しました！（船坂昭美 除外済）")
