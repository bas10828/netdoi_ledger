"""Export slip_transactions ledger to Excel + PDF for management review.

Run: python export_report.py
Output: ledger_report.xlsx, ledger_report.pdf
"""

import os

import pandas as pd
import psycopg2
from dotenv import load_dotenv
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

load_dotenv()

FONT = "Tahoma"
pdfmetrics.registerFont(TTFont(FONT, r"C:\Windows\Fonts\tahoma.ttf"))

OUT_XLSX = "ledger_report.xlsx"
OUT_PDF = "ledger_report.pdf"

QUERY = """
SELECT
    txn_date AS "วันที่",
    txn_time AS "เวลา",
    bank AS "ธนาคาร",
    direction AS "ประเภท",
    amount AS "ยอดเงิน",
    fee AS "ค่าธรรมเนียม",
    sender_name AS "ผู้โอน",
    receiver_name AS "ผู้รับ",
    memo AS "รายละเอียด"
FROM slip_transactions
ORDER BY txn_date, txn_time
"""


def fetch_df():
    conn = psycopg2.connect(os.environ["DATABASE_URL"])
    try:
        return pd.read_sql(QUERY, conn)
    finally:
        conn.close()


def write_excel(df: pd.DataFrame):
    summary = (
        df.groupby("ประเภท")["ยอดเงิน"]
        .sum()
        .reindex(["expense", "income", "unknown"])
        .fillna(0)
        .rename({"expense": "รายจ่าย", "income": "รายรับ", "unknown": "ไม่ระบุ"})
    )
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="รายการ", index=False)
        summary.to_frame("รวม (บาท)").to_excel(writer, sheet_name="สรุป")
    print(f"-> {OUT_XLSX}")


def write_pdf(df: pd.DataFrame):
    styles = getSampleStyleSheet()
    h1 = ParagraphStyle("h1", parent=styles["Title"], fontName=FONT, fontSize=16, spaceAfter=10)
    body = ParagraphStyle("body", parent=styles["BodyText"], fontName=FONT, fontSize=10, leading=14)

    header = list(df.columns)
    rows = df.fillna("").astype(str).values.tolist()
    table_data = [header] + rows

    t = Table(table_data, repeatRows=1)
    t.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), FONT),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f4f6f7")]),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))

    total_expense = df.loc[df["ประเภท"] == "expense", "ยอดเงิน"].sum()
    total_income = df.loc[df["ประเภท"] == "income", "ยอดเงิน"].sum()

    story = [
        Paragraph("รายงานสรุปรายการเงิน (สลิปธนาคาร)", h1),
        Paragraph(f"จำนวนรายการทั้งหมด: {len(df)} รายการ", body),
        Paragraph(f"รวมรายจ่าย: {total_expense:,.2f} บาท &nbsp;&nbsp; รวมรายรับ: {total_income:,.2f} บาท", body),
        Spacer(1, 0.4 * cm),
        t,
    ]
    SimpleDocTemplate(
        OUT_PDF, pagesize=landscape(A4),
        topMargin=1.5 * cm, bottomMargin=1.5 * cm, leftMargin=1.2 * cm, rightMargin=1.2 * cm,
    ).build(story)
    print(f"-> {OUT_PDF}")


def main():
    df = fetch_df()
    write_excel(df)
    write_pdf(df)


if __name__ == "__main__":
    main()
