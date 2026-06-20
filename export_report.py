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
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.shapes import Drawing, Rect, String

load_dotenv()

PALETTE = [
    "#c0392b", "#e67e22", "#f1c40f", "#2c70c9", "#5a4fcf", "#1e8449",
    "#97a3ad", "#8e44ad", "#16a085", "#d35400", "#7f8c8d", "#2980b9",
]

FONT = "ThaiReportFont"
_FONT_CANDIDATES = [
    r"C:\Windows\Fonts\tahoma.ttf",
    "/usr/share/fonts/truetype/tlwg/Sarabun.ttf",
    "/usr/share/fonts/truetype/thai-tlwg/Sarabun.ttf",
    "/usr/share/fonts/truetype/noto/NotoSansThai-Regular.ttf",
]
_font_path = next((p for p in _FONT_CANDIDATES if os.path.exists(p)), None)
if not _font_path:
    raise RuntimeError(
        "No Thai-capable TTF font found (checked: "
        + ", ".join(_FONT_CANDIDATES)
        + "). Install a Thai font package (e.g. fonts-thai-tlwg) on this host/container."
    )
pdfmetrics.registerFont(TTFont(FONT, _font_path))

OUT_XLSX = "ledger_report.xlsx"
OUT_PDF = "ledger_report.pdf"

QUERY = """
SELECT
    txn_date AS "วันที่",
    txn_time AS "เวลา",
    bank AS "ธนาคาร",
    direction AS "ประเภท",
    category AS "หมวด",
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


def build_pdf(df: pd.DataFrame, output):
    """Render df to a PDF. `output` is a file path (str) or a file-like object (e.g. BytesIO)."""
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
        output, pagesize=landscape(A4),
        topMargin=1.5 * cm, bottomMargin=1.5 * cm, leftMargin=1.2 * cm, rightMargin=1.2 * cm,
    ).build(story)


def _bar_chart_drawing(labels, expense, income):
    width, height = 480, 220
    d = Drawing(width, height)
    chart = VerticalBarChart()
    chart.x = 45
    chart.y = 45
    chart.width = width - 70
    chart.height = height - 80
    chart.data = [expense, income]
    chart.categoryAxis.categoryNames = labels
    chart.categoryAxis.labels.fontName = FONT
    chart.categoryAxis.labels.fontSize = 7
    chart.categoryAxis.labels.angle = 30
    chart.categoryAxis.labels.dy = -10
    chart.valueAxis.labels.fontName = FONT
    chart.valueAxis.labels.fontSize = 7
    chart.bars[0].fillColor = colors.HexColor("#c0392b")
    chart.bars[1].fillColor = colors.HexColor("#1e8449")
    chart.groupSpacing = 10
    d.add(chart)
    d.add(Rect(width - 150, height - 14, 9, 9, fillColor=colors.HexColor("#c0392b"), strokeColor=None))
    d.add(String(width - 137, height - 14, "รายจ่าย", fontName=FONT, fontSize=8))
    d.add(Rect(width - 70, height - 14, 9, 9, fillColor=colors.HexColor("#1e8449"), strokeColor=None))
    d.add(String(width - 57, height - 14, "รายรับ", fontName=FONT, fontSize=8))
    return d


def _pie_chart_drawing(category_rows):
    if not category_rows:
        return None
    row_height = 20
    height = max(170, row_height * len(category_rows) + 20)
    d = Drawing(480, height)
    pie = Pie()
    pie.x = 30
    pie.y = (height - 150) / 2
    pie.width = 150
    pie.height = 150
    pie.data = [c["total"] for c in category_rows]
    for i in range(len(category_rows)):
        pie.slices[i].fillColor = colors.HexColor(PALETTE[i % len(PALETTE)])
        pie.slices[i].strokeColor = colors.white
        pie.slices[i].strokeWidth = 1
    d.add(pie)
    legend_x = 220
    for i, c in enumerate(category_rows):
        y = height - 18 - i * row_height
        d.add(Rect(legend_x, y, 10, 10, fillColor=colors.HexColor(PALETTE[i % len(PALETTE)]), strokeColor=None))
        d.add(String(legend_x + 15, y + 1, f"{c['name']} ({c['total']:,.0f})", fontName=FONT, fontSize=8))
    return d


def build_summary_pdf(period_label, total_expense, total_income, net, category_rows, top_payees, bucket_chart, output):
    """Management summary (totals + bar/pie charts + category breakdown + top payees) for a single period."""
    styles = getSampleStyleSheet()
    h1 = ParagraphStyle("h1", parent=styles["Title"], fontName=FONT, fontSize=16, spaceAfter=10)
    h2 = ParagraphStyle("h2", parent=styles["Heading2"], fontName=FONT, fontSize=12, spaceBefore=16, spaceAfter=6)
    body = ParagraphStyle("body", parent=styles["BodyText"], fontName=FONT, fontSize=11, leading=16)

    def table(header, rows):
        t = Table([header] + rows, repeatRows=1)
        t.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), FONT),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f4f6f7")]),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        return t

    story = [
        Paragraph(f"รายงานสรุป — {period_label}", h1),
        Paragraph(
            f"รายจ่าย: {total_expense:,.2f} บาท &nbsp;&nbsp; "
            f"รายรับ: {total_income:,.2f} บาท &nbsp;&nbsp; "
            f"สุทธิ: {net:,.2f} บาท",
            body,
        ),
    ]

    if bucket_chart["labels"]:
        story.append(Paragraph("รายรับ-รายจ่ายตามช่วงเวลา", h2))
        story.append(_bar_chart_drawing(bucket_chart["labels"], bucket_chart["expense"], bucket_chart["income"]))

    pie = _pie_chart_drawing(category_rows)
    if pie:
        story.append(Paragraph("แยกตามหมวด", h2))
        story.append(pie)

    story.append(Paragraph("แยกตามหมวด (ตัวเลข)", h2))
    story.append(table(
        ["หมวด", "ยอด (บาท)"],
        [[c["name"], f'{c["total"]:,.2f}'] for c in category_rows] or [["-", "-"]],
    ))
    story.append(Paragraph("Top ผู้รับเงิน", h2))
    story.append(table(
        ["อันดับ", "ผู้รับ", "ยอด (บาท)"],
        [[i + 1, p["name"], f'{p["total"]:,.2f}'] for i, p in enumerate(top_payees)] or [["-", "-", "-"]],
    ))

    SimpleDocTemplate(
        output, pagesize=A4,
        topMargin=2 * cm, bottomMargin=2 * cm, leftMargin=2 * cm, rightMargin=2 * cm,
    ).build(story)


def main():
    df = fetch_df()
    write_excel(df)
    build_pdf(df, OUT_PDF)
    print(f"-> {OUT_PDF}")


if __name__ == "__main__":
    main()
