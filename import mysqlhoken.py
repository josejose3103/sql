import mysql.connector
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import datetime

# ユーザー入力
start_day = input("開始日を入力（YYYY-MM-DD）: ")
end_day = input("終了日を入力（YYYY-MM-DD）: ")
kumikan = input("組勘番号を入力: ")

# MySQLデータベース接続
conn = mysql.connector.connect(
    host="localhost",      # MySQLサーバーのホスト名
    user="root",           # MySQLのユーザー名
    password="3103@Kazu",  # MySQLのパスワード（セキュリティに注意）
    database="media",      # 使用するデータベース名
    charset="utf8"
)
cursor = conn.cursor()

# SQL実行
query = """
 SELECT day, noukamei, kingaku, byoumei
 FROM article1 
 WHERE day BETWEEN %s AND %s AND kumikan = %s; 
"""
cursor.execute(query, (start_day, end_day, kumikan))

# 結果を取得
rows = cursor.fetchall()

# PDFファイル名（現在の日付を含める）
pdf_filename = f"出力結果_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
# 日本語フォントの登録（IPAexゴシックを使用）



# 日本語フォントの登録（IPAexゴシックを使用）
#pdfmetrics.registerFont(TTFont("IPAexGothic", "/usr/share/fonts/opentype/ipaexg.ttf"))

# PDFを作成
pdf = canvas.Canvas(pdf_filename, pagesize=A4)


pdfmetrics.registerFont(TTFont("MSGothic", "C:/Windows/Fonts/msgothic.ttc"))
# フォント設定（日本語対応）
pdf.setFont("MSGothic", 12)
# フォント設定（日本語対応）
#pdf.setFont("IPAexGothic", 12)

# タイトル
pdf.drawString(100, 800, "検索結果")

# ヘッダー
pdf.drawString(100, 780, "番号    病名    金額")

# データを書き込み
y_position = 760  # 初期Y位置
for row in rows:
    number, byoumei, kingaku, byoumei = row
    pdf.drawString(100, y_position, f"{number}    {byoumei}    {kingaku}円")
    y_position -= 20  # 1行下げる

# PDFを保存
pdf.save()

# 終了処理
cursor.close()
conn.close()

print(f"PDFファイル '{pdf_filename}' を作成しました。")
