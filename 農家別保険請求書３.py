import pyodbc
import pandas as pd
import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import os

# ✅ 日本語フォントを登録
pdfmetrics.registerFont(TTFont('MSGothic', 'C:/Windows/Fonts/msgothic.ttc'))  # Windows用

# ユーザーに日付・組勘番号・農家名・請求期間を入力
start_date = input("開始日を YYYY/MM/DD 形式で入力してください: ")
end_date = input("終了日を YYYY/MM/DD 形式で入力してください: ")
group_number = input("組勘番号を入力してください: ")
farmer_name = input("農家名を入力してください: ")
billing_period = input("請求期間 (例: 1月1日から1月31日まで) を入力してください: ")

# SQL の日付フォーマットに合わせる
start_date_sql = f"#{start_date}#"
end_date_sql = f"#{end_date}#"

conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\josej\OneDrive\Documents\カルテ.mdb;'
)
conn = pyodbc.connect(conn_str)

# SQLクエリを実行
sql = f"""
    SELECT 番号, 病名, 金額 
    FROM 平成１５年 
    WHERE 日付 BETWEEN {start_date_sql} AND {end_date_sql} 
    AND 組間番号 = {group_number}
    ORDER BY 番号 ASC
"""

df = pd.read_sql(sql, conn)

# ✅ 内税方式の計算（3桁カンマ & ¥付き）
total_amount = df["金額"].sum()
tax_excluded = round(total_amount / 1.1)
tax = total_amount - tax_excluded

def format_currency(value):
    return f"¥{value:,.0f}"

total_amount_f = format_currency(total_amount)
tax_excluded_f = format_currency(tax_excluded)
tax_f = format_currency(tax)

# 今日の日付
today = datetime.date.today().strftime("%Y%m%d")
save_dir = r"C:\Users\josej\OneDrive\Documents\pdf領収書"
pdf_filename = f"{farmer_name}_{today}.pdf"
c = canvas.Canvas(pdf_filename, pagesize=A4)
width, height = A4

def draw_header():
    c.setFont("MSGothic", 20)
    c.drawCentredString(width / 2, height - 50, "請求書")
    c.setFont("MSGothic", 12)
    c.drawCentredString(width / 2, height - 90, "弟子屈町奥春別")
    c.drawCentredString(width / 2, height - 110, f"{farmer_name} 様")
    c.drawCentredString(width / 2, height - 130, f"処理日: {today}")
    c.drawString(50, height - 160, "番号")
    c.drawString(150, height - 160, "病名")
    c.drawString(350, height - 160, "金額（税込）")
    c.line(50, height - 165, 500, height - 165)

# ✅ ヘッダー描画
draw_header()

# データ描画（ページまたぎ対応）
y_position = height - 190
for index, row in df.iterrows():
    if y_position < 100:
        c.showPage()
        draw_header()
        y_position = height - 190

    formatted_amount = format_currency(row["金額"])
    c.drawString(50, y_position, str(row["番号"]))
    c.drawString(150, y_position, row["病名"])
    c.drawString(350, y_position, formatted_amount)
    y_position -= 20

# ✅ 金額集計欄（改ページ判定）
if y_position < 150:
    c.showPage()
    c.setFont("MSGothic", 12)
    y_position = height - 50

# ✅ 合計金額部分
c.drawString(50, y_position, "税抜き合計")
c.drawString(350, y_position, tax_excluded_f)
y_position -= 20

c.drawString(50, y_position, "消費税 (10%)")
c.drawString(350, y_position, tax_f)
y_position -= 20

c.drawString(50, y_position, "総合計（税込）")
c.drawString(350, y_position, total_amount_f)
y_position -= 40

c.setFont("MSGothic", 24)
c.drawCentredString(width / 2, y_position, "ご請求金額")
y_position -= 30
c.drawCentredString(width / 2, y_position, total_amount_f)
y_position -= 40

# ✅ 会社情報・振込口座情報
c.setFont("MSGothic", 12)
c.drawRightString(width - 50, y_position, "北海道川上郡弟子屈町朝日2丁目4-30")
y_position -= 15
c.drawRightString(width - 50, y_position, "株式会社 TUB （T9560001004706）")
y_position -= 15
c.drawRightString(width - 50, y_position, "代表取締役 坪井 聡之")
y_position -= 20
c.drawRightString(width - 50, y_position, "下記の口座に入金してください")
y_position -= 15
c.drawRightString(width - 50, y_position, "摩周湖農業協同組合　本所")
y_position -= 15
c.drawRightString(width - 50, y_position, "普通　0022095")
y_position -= 15
c.drawRightString(width - 50, y_position, "株式会社 TUB")
y_position -= 15
c.drawRightString(width - 50, y_position, "代表取締役 坪井 聡之")

# PDF保存
c.save()

# 保存先ディレクトリに移動
if not os.path.exists(save_dir):
    os.makedirs(save_dir)
pdf_path = os.path.join(save_dir, pdf_filename)
os.rename(pdf_filename, pdf_path)

print(f"PDFファイル '{pdf_filename}' を作成しました。")
input("Press Enter to exit...")
