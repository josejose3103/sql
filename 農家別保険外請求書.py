import pyodbc
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

# 日本語フォントの登録
pdfmetrics.registerFont(TTFont('MSGothic', 'msgothic.ttc'))

# ユーザー入力
start_date = input("開始日 (YYYY/MM/DD): ")
end_date = input("終了日 (YYYY/MM/DD): ")
address = input("住所: ")
field_4_value = input("フィールド４の値: ")
process_date = pd.Timestamp.now().strftime('%Y-%m-%d')

conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\\Users\\josej\\OneDrive\\Documents\\カルテ.mdb;'
)
conn = pyodbc.connect(conn_str)

sql = f"""
SELECT フィールド1 AS 月日, フィールド2 AS 診療損害防止等の内容, フィールド3 AS 金額
FROM 保険外　薬品
WHERE フィールド1 BETWEEN #{start_date}# AND #{end_date}#
AND フィールド４ = '{field_4_value}'
"""

df = pd.read_sql(sql, conn)
df.columns = ['月日', '診療・損害防止等の内容', '金額']
df['月日'] = pd.to_datetime(df['月日']).dt.strftime('%Y-%m-%d')
df['金額'] = df['金額'].astype(int)
conn.close()

# 内税方式の計算
total_with_tax = df['金額'].sum()
tax_amount = int(total_with_tax / 11)
total_amount = total_with_tax - tax_amount

def export_to_pdf(df, filename="保険外請求書.pdf"):
    doc = SimpleDocTemplate(filename, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()
    right_align = ParagraphStyle(name="RightAlign", alignment=2, fontName='MSGothic', fontSize=12)
    
    # タイトル
    title = Paragraph("<font name='MSGothic' size='18'><b>請求書</b></font>", styles['Title'])
    elements.append(title)
    elements.append(Spacer(1, 12))
    
    # 住所
    address_paragraph = Paragraph(f"<font name='MSGothic' size='12'>{address}</font>", styles['Normal'])
    elements.append(address_paragraph)
    elements.append(Spacer(1, 12))
    
    # フィールド4の値 + 様
    field_4_paragraph = Paragraph(f"<font name='MSGothic' size='12'>{field_4_value} 様</font>", styles['Normal'])
    elements.append(field_4_paragraph)
    elements.append(Spacer(1, 12))
    
    # 処理日
    process_date_paragraph = Paragraph(f"<font name='MSGothic' size='12'>処理日: {process_date}</font>", styles['Normal'])
    elements.append(process_date_paragraph)
    elements.append(Spacer(1, 20))
    
    # テーブルデータ
    data = [df.columns.tolist()] + df.values.tolist()
    table = Table(data)
    
    # テーブルスタイル
    style = TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'MSGothic'),
        ('BACKGROUND', (0, 0), (-1, 0), colors.gray),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ])
    table.setStyle(style)
    
    elements.append(table)
    elements.append(Spacer(1, 20))
    
    # 合計情報（表とは別）
    total_paragraph = Paragraph(f"<font name='MSGothic' size='12'>税抜金額: ¥{total_amount:,}</font>", styles['Normal'])
    elements.append(total_paragraph)
    elements.append(Spacer(1, 12))
    
    tax_paragraph = Paragraph(f"<font name='MSGothic' size='12'>内税消費税額: ¥{tax_amount:,}</font>", styles['Normal'])
    elements.append(tax_paragraph)
    elements.append(Spacer(1, 12))
    
    total_with_tax_paragraph = Paragraph(f"<font name='MSGothic' size='16'><b>ご請求金額: ¥{total_with_tax:,}</b></font>", styles['Normal'])
    elements.append(total_with_tax_paragraph)
    elements.append(Spacer(1, 20))
    
    # 会社情報（右寄せ）
    company_info = Paragraph("北海道川上郡弟子屈町朝日2丁目4-30<br/>株式会社 TUB (T9560001004706)<br/>獣医師 坪井 聡之", right_align)
    elements.append(company_info)
    elements.append(Spacer(1, 12))
    
    # 振込情報（右寄せ）
    payment_info = Paragraph("下記の口座に入金してください<br/>摩周湖農業協同組合 本所<br/>普通 0022095<br/>株式会社 TUB<br/>代表取締役 坪井 聡之", right_align)
    elements.append(payment_info)
    
    doc.build(elements)

# PDFファイルに保存
export_to_pdf(df)
print("PDFファイル '保険外請求書.pdf' を作成しました。")
