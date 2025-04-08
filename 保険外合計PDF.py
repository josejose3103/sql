import pyodbc
import pandas as pd
from reportlab.lib.pagesizes import A4, portrait
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime
import os

# 日本語フォントの登録（IPAexフォントを使用）
# フォントをダウンロードして配置したパスを指定
font_path = "C:/Users/josej/AppData/Local/Microsoft/Windows/Fonts/ipaexg.ttf"  # このパスは実際のフォントファイルの場所に変更してください

# フォントが見つからない場合のエラーメッセージ
if not os.path.exists(font_path):
    print(f"エラー: フォントファイル {font_path} が見つかりません。")
    print("IPAフォントをダウンロードして指定のパスに配置してください。")
    print("ダウンロード先: https://moji.or.jp/ipafont/")
    exit(1)

# フォントの登録
pdfmetrics.registerFont(TTFont('IPAexGothic', font_path))

# データベース接続
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\josej\OneDrive\Documents\カルテ.mdb;'
)
conn = pyodbc.connect(conn_str)

# SQLクエリ実行
sql = "SELECT フィールド４, sum(フィールド3) FROM 保険外　薬品 WHERE フィールド1 BETWEEN #2025/04/01# AND #2025/05/31# GROUP BY フィールド４"
df = pd.read_sql(sql, conn)

# 列名を設定
df.columns = ['農家名', '合計金額']

# PDFファイル名を設定
current_date = datetime.now().strftime("%Y%m%d")
pdf_filename = f"保険外合計請求書_{current_date}.pdf"

# ReportLabでPDF作成
doc = SimpleDocTemplate(pdf_filename, pagesize=portrait(A4))
elements = []

# 日本語スタイルの設定
styles = getSampleStyleSheet()
title_style = ParagraphStyle(
    name='JapaneseTitle',
    parent=styles['Title'],
    fontName='IPAexGothic',
    fontSize=16
)
normal_style = ParagraphStyle(
    name='JapaneseNormal',
    parent=styles['Normal'],
    fontName='IPAexGothic',
    fontSize=12
)

# タイトルと期間
elements.append(Paragraph("保険外合計請求書", title_style))
elements.append(Paragraph("期間: 2025年4月1日 ～ 2025年5月31日", normal_style))
elements.append(Spacer(1, 20))

# 「船坂昭美」と「渋谷敦士」を除外
df = df[~df["農家名"].isin(["船坂昭美", "渋谷敦士"])]

# 金額を「￥3,000」形式に変更
df["合計金額"] = df["合計金額"].apply(lambda x: f"¥{x:,.0f}")

# データをテーブル用に準備
data = [['農家名', '合計金額']]
total = 0
for i, row in df.iterrows():
    data.append([str(row["農家名"]), row["合計金額"]])
    
    # 数値に戻して合計計算（「¥」「,」を削除）
    total += int(row["合計金額"].replace("¥", "").replace(",", ""))

# 合計金額も「￥3,000」形式に変更
data.append(['合計', f"¥{total:,.0f}"])


# テーブル作成
table = Table(data, colWidths=[270, 270])
table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (1, 0), colors.grey),
    ('TEXTCOLOR', (0, 0), (1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (1, -1), 'CENTER'),
    ('FONTNAME', (0, 0), (1, -1), 'IPAexGothic'),  # テーブル内の日本語フォント
    ('BOTTOMPADDING', (0, 0), (1, 0), 12),
    ('BACKGROUND', (0, -1), (1, -1), colors.lightgrey),
    ('GRID', (0, 0), (1, -1), 1, colors.black),
]))

elements.append(table)

# PDF生成
doc.build(elements)
print(f"PDFファイルが正常に作成されました: {pdf_filename}")
# 保存先ディレクトリを設定
output_dir = "C:/Users/josej/OneDrive/Documents/pdf領収書"

# ディレクトリが存在しない場合は作成
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# PDFファイルを保存先ディレクトリに移動
output_path = os.path.join(output_dir, pdf_filename)
os.rename(pdf_filename, output_path)

print(f"PDFファイルが正常に保存されました: {output_path}")
# 接続を閉じる
conn.close()