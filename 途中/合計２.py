import pyodbc
import pandas as pd
from fpdf import FPDF

# データベース接続情報
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\josej\OneDrive\Documents\カルテ.mdb;'
)
conn = pyodbc.connect(conn_str)

# ユーザー入力で日付を指定
start_date = input("開始日 (YYYY/MM/DD): ")
end_date = input("終了日 (YYYY/MM/DD): ")

# SQL クエリの動的生成
sql = "SELECT 農家名, sum(金額) as 合計金額 FROM 平成１５年 WHERE 日付 BETWEEN #{start_date}# AND #{end_date}# GROUP BY 農家名"

# データ取得
df = pd.read_sql(sql, conn)

# PDF出力の設定
pdf = FPDF()
pdf.add_page()

# 日本語フォントを設定 (MS Gothic など)
pdf.add_font('MSGothic', '', r'C:\Windows\Fonts\msgothic.ttc', uni=True)
pdf.set_font("MSGothic", size=12)

# タイトル
pdf.cell(200, 10, txt="集計結果", ln=True, align="C")
pdf.cell(200, 10, txt=f"期間: {start_date} 〜 {end_date}", ln=True, align="C")

# ヘッダー
pdf.cell(100, 10, txt="農家名", border=1, align="C")
pdf.cell(50, 10, txt="合計金額", border=1, align="C")
pdf.ln()

# データ追加
for index, row in df.iterrows():
    pdf.cell(100, 10, txt=str(row["農家名"]), border=1)
    pdf.cell(50, 10, txt=str(row["合計金額"]), border=1, align="R")
    pdf.ln()

# PDFを保存
pdf_output_path = "集計結果.pdf"
pdf.output(pdf_output_path)
print(f"PDFを作成しました: {pdf_output_path}")
