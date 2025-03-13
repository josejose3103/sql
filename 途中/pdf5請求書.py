import pyodbc
import pandas as pd
import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

# ✅ 日本語フォントを登録
pdfmetrics.registerFont(TTFont('MSGothic', 'C:/Windows/Fonts/msgothic.ttc'))  # Windows用
# pdfmetrics.registerFont(TTFont('IPAexGothic', '/usr/share/fonts/opentype/ipafont-mincho/ipag.ttf'))  # Mac/Linux用

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

# ✅ 内税方式の計算
total_amount = df["金額"].sum()  # 総額（税込）
tax_excluded = round(total_amount / 1.1)  # 税抜き金額
tax = total_amount - tax_excluded  # 消費税額

# 今日の日付
today = datetime.date.today().strftime("%Y/%m/%d")

# ✅ PDFを作成
pdf_filename = "invoice.pdf"
c = canvas.Canvas(pdf_filename, pagesize=A4)
width, height = A4

# ✅ タイトル（中央配置・フォント大きめ）
c.setFont("MSGothic", 20)  
c.drawCentredString(width / 2, height - 50, "請求書")

# ✅ 農家名（中央揃え）
c.setFont("MSGothic", 12)
c.drawCentredString(width / 2, height - 90, "弟子屈町奥春別")  # 農家名の上の行
c.drawCentredString(width / 2, height - 110, f"{farmer_name} 様")  # 「様」を付加

# ✅ 日付（中央揃え）
c.drawCentredString(width / 2, height - 130, f"処理日: {today}")

# ✅ ヘッダー（番号と病名の間隔を広げる）
c.drawString(50, height - 160, "番号")
c.drawString(150, height - 160, "病名")  # 100 → 150px で広げた
c.drawString(350, height - 160, "金額（税込）")

# 横線
c.line(50, height - 165, 500, height - 165)

# ✅ データ部分（病名の位置を右へ移動）
y_position = height - 190
for index, row in df.iterrows():
    c.drawString(50, y_position, str(row["番号"]))  # 番号（左寄せ）
    c.drawString(150, y_position, row["病名"])  # 病名（間隔広げた）
    c.drawString(350, y_position, f"{row['金額']} 円")  # 金額
    y_position -= 20

# ✅ 内税方式の金額を表示
y_position -= 30
c.drawString(50, y_position, "税抜き合計")
c.drawString(350, y_position, f"{tax_excluded} 円")

y_position -= 20
c.drawString(50, y_position, "消費税 (10%)")
c.drawString(350, y_position, f"{tax} 円")

y_position -= 20
c.drawString(50, y_position, "総合計（税込）")
c.drawString(350, y_position, f"{total_amount} 円")

# ✅ 「総合計 を請求いたします」と請求期間
y_position -= 30
c.drawString(50, y_position, "総合計 を請求いたします")
c.drawRightString(width - 50, y_position, f"（{billing_period}）")

# ✅ 「ご請求金額」と「総合計（税込）」を**中央寄せ** & **フォント2倍**
y_position -= 40
c.setFont("MSGothic", 24)  # フォントを2倍
c.drawCentredString(width / 2, y_position, "ご請求金額")

y_position -= 30
c.drawCentredString(width / 2, y_position, f"{total_amount} 円")

# ✅ 「ご利用いただきありがとうございます。」を中央に追加
c.setFont("MSGothic", 14)
c.drawCentredString(width / 2, 100, "ご利用いただきありがとうございます。")

# ✅ 会社情報を右寄せで配置
c.setFont("MSGothic", 10)
c.drawRightString(width - 50, 80, "北海道川上郡弟子屈町朝日2丁目4-30")
c.drawRightString(width - 50, 65, "株式会社 TUB （T9560001004706）")
c.drawRightString(width - 50, 50, "獣医師 坪井 聡之")

# ✅ 振込口座情報を右寄せで配置
c.setFont("MSGothic", 10)
c.drawRightString(width - 50, 30, "下記の口座に入金してください")
c.drawRightString(width - 50, 15, "摩周湖農業協同組合　本所")
c.drawRightString(width - 50, 0, "普通　0022095")
c.drawRightString(width - 50, -15, "株式会社 TUB")

# PDFを保存
c.save()

# ✅ 結果を表示
print(f"PDFファイル '{pdf_filename}' を作成しました。")
print(f"農家名: {farmer_name} 様")
print(f"請求期間: {billing_period}")
print(f"処理日: {today}")
print(f"税抜き合計: {tax_excluded} 円")
print(f"消費税 (10%): {tax} 円")
print(f"総合計（税込）: {total_amount} 円")
