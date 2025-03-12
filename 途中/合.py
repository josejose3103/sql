import pyodbc
import pandas as pd
from fpdf import FPDF

# Database connection information
try:
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=C:\Users\josej\OneDrive\Documents\カルテ.mdb;'
    )
    conn = pyodbc.connect(conn_str)
except Exception as e:
    print(f"Connection error: {e}")
    exit()

# User input for start and end dates
try:
    start_date = input("開始日 (YYYY/MM/DD): ")
    end_date = input("終了日 (YYYY/MM/DD): ")
except ValueError:
    print("Invalid date format. Please use YYYY/MM/DD.")
    exit()

# SQL query using parameters to avoid formatting issues
sql = """
SELECT 農家名, sum(金額) as total_amount 
FROM 平成１５年 
WHERE 日付 BETWEEN ? AND ?
GROUP BY 農家名
"""
try:
    df = pd.read_sql(sql, conn, params=[start_date, end_date])
except Exception as e:
    print(f"SQL error: {e}")
    exit()

# PDF output settings
pdf = FPDF()
pdf.add_page()
pdf.set_font("Arial", size=12)

# Title and period information
pdf.cell(200, 10, txt="集計結果", ln=True, align="C")
pdf.cell(200, 10, txt=f"期間: {start_date} 〜 {end_date}", ln=True, align="C")

# Header with right alignment for numeric values
header = [
    ("農家名", 100),
    ("合計金額", 50)
]
for col, width in header:
    pdf.cell(width, 10, txt=col, border=1, align="C")
pdf.ln()

# Data output with proper formatting
try:
    for index, row in df.iterrows():
        farmer_name = str(row["農家名"])
        amount = float(row["total_amount"])
        # Format the amount to avoid scientific notation for small numbers
        if 0 < amount < 100000:
            amount_str = f"{amount:.2f}"
        else:
            amount_str = f"{int(amount)}" if int(amount) == amount else "{:.6f}".format(amount)
        
        pdf.cell(100, 10, txt=farmer_name, border=1)
        pdf.cell(50, 10, txt=amount_str, border=1, align="R")
        pdf.ln()
except Exception as e:
    print(f"Data processing error: {e}")
    exit()

# Save PDF and print confirmation
pdf_output_path = "集計結果.pdf"
pdf.output(pdf_output_path)
print(f"PDFを作成しました: {pdf_output_path}")
