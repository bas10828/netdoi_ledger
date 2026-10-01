"""Expense report export (Excel + PDF) for a date range, plus the plain ledger-listing PDF.

fetch_period_report() builds one report dict that both build_report_xlsx() and
build_report_pdf() render, so the two files always carry the same numbers.

Run: python export_report.py [YYYY-MM]
Output: report_YYYY-MM.xlsx, report_YYYY-MM.pdf (defaults to the current month)
"""

import glob
import os
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta, timezone
from xml.sax.saxutils import escape

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table as XLTable, TableStyleInfo
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.shapes import Circle, Drawing, Line, Rect, String
from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import registerFontFamily, stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.platypus import (
    CondPageBreak, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

from categorize import get_categories

# ---------------------------------------------------------------------------
# Fonts
# ---------------------------------------------------------------------------

FONT = "ThaiReportFont"
FONT_BOLD = "ThaiReportFont-Bold"
_FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
# Bundled Sarabun (OFL) first: it covers Thai *and* Latin/digits/punctuation. Debian's
# NotoSansThai is Thai-only (no digits even), so it must not be used on its own here.
_FONT_CANDIDATES = [
    (os.path.join(_FONT_DIR, "Sarabun-Regular.ttf"), os.path.join(_FONT_DIR, "Sarabun-Bold.ttf")),
    ("/usr/share/fonts/truetype/tlwg/Sarabun.ttf", "/usr/share/fonts/truetype/tlwg/Sarabun-Bold.ttf"),
    ("/usr/share/fonts/truetype/thai-tlwg/Sarabun.ttf", "/usr/share/fonts/truetype/thai-tlwg/Sarabun-Bold.ttf"),
    (r"C:\Windows\Fonts\tahoma.ttf", r"C:\Windows\Fonts\tahomabd.ttf"),
]
_FONT_GLOBS = [
    ("/usr/share/fonts/**/*Sarabun*.ttf", "/usr/share/fonts/**/*Sarabun*Bold*.ttf"),
]


def _find_fonts():
    for regular, bold in _FONT_CANDIDATES:
        if os.path.exists(regular):
            return regular, bold if os.path.exists(bold) else None
    for reg_pattern, bold_pattern in _FONT_GLOBS:
        regular = next(iter(sorted(glob.glob(reg_pattern, recursive=True))), None)
        if regular:
            bold = next(iter(sorted(glob.glob(bold_pattern, recursive=True))), None) if bold_pattern else None
            return regular, bold
    return None, None


_font_path, _bold_path = _find_fonts()
if not _font_path:
    raise RuntimeError(
        "No Thai-capable TTF font found (checked: "
        + ", ".join(p for pair in _FONT_CANDIDATES + _FONT_GLOBS for p in pair if p)
        + "). The bundled fonts/ directory is missing from this host/container."
    )
pdfmetrics.registerFont(TTFont(FONT, _font_path))
# No bold face available -> bold text silently falls back to the regular face.
pdfmetrics.registerFont(TTFont(FONT_BOLD, _bold_path or _font_path))
registerFontFamily(FONT, normal=FONT, bold=FONT_BOLD, italic=FONT, boldItalic=FONT_BOLD)

# Thai stacks tone marks over upper vowels (ที่, น้ำ, ตั้ง); without HarfBuzz shaping they collide.
try:
    import uharfbuzz  # noqa: F401
    SHAPING = 1
except ImportError:
    SHAPING = 0


def _ps(name, **kw):
    return ParagraphStyle(name, shaping=SHAPING, **kw)

# ---------------------------------------------------------------------------
# Shared constants (kept in step with the web theme in templates/base.html)
# ---------------------------------------------------------------------------

BANGKOK_TZ = timezone(timedelta(hours=7))
UNCATEGORIZED = "ไม่ระบุหมวด"
UNKNOWN_PAYEE = "ไม่ระบุผู้รับ"

THAI_MONTHS = {
    1: "ม.ค.", 2: "ก.พ.", 3: "มี.ค.", 4: "เม.ย.", 5: "พ.ค.", 6: "มิ.ย.",
    7: "ก.ค.", 8: "ส.ค.", 9: "ก.ย.", 10: "ต.ค.", 11: "พ.ย.", 12: "ธ.ค.",
}
THAI_MONTHS_FULL = {
    1: "มกราคม", 2: "กุมภาพันธ์", 3: "มีนาคม", 4: "เมษายน", 5: "พฤษภาคม", 6: "มิถุนายน",
    7: "กรกฎาคม", 8: "สิงหาคม", 9: "กันยายน", 10: "ตุลาคม", 11: "พฤศจิกายน", 12: "ธันวาคม",
}

C_INK = "111827"
C_MUTED = "6B7280"
C_FAINT = "9CA3AF"
C_BORDER = "E5E7EB"
C_SURFACE_ALT = "F8FAFC"
C_HEADER_BAND = "1F2937"
C_ACCENT = "2F5BD3"
C_ACCENT_SOFT = "EEF2FD"
C_UP = "C8372D"        # spending went up -> bad
C_DOWN = "15803D"      # spending went down -> good
C_WARN = "B45309"
C_WARN_SOFT = "FEF3E2"

# Same fixed-order categorical palette as the /reports page; anything past it is "other".
PALETTE = ["2A78D6", "1BAF7A", "EDA100", "008300", "4A3AA7", "E34948", "E87BA4", "EB6834"]
OTHER_COLOR = "B4BCC5"
PIE_SLICES = 7

REPORT_TITLE = "รายงานรายจ่าย"


def _hex(c):
    return colors.HexColor("#" + c)


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------

def fmt_baht(v):
    return f"{v:,.2f}"


def fmt_pct(v, signed=False):
    if v is None:
        return "–"
    return f"{v:+.1f}%" if signed else f"{v:.1f}%"


def fmt_date(d, year=True):
    s = f"{d.day} {THAI_MONTHS[d.month]}"
    return f"{s} {d.year}" if year else s


def pct_change(cur, prev):
    if not prev:
        return None
    return (cur - prev) / prev * 100


def _month_end(d):
    nxt = date(d.year + 1, 1, 1) if d.month == 12 else date(d.year, d.month + 1, 1)
    return nxt - timedelta(days=1)


# ---------------------------------------------------------------------------
# Data
# ---------------------------------------------------------------------------

def fetch_period_report(cur, d_from, d_to, today=None):
    """Everything the expense report shows for [d_from, d_to], compared against the
    previous period: the prior calendar month when the range is exactly one month,
    otherwise the equally long window right before it."""
    today = today or datetime.now(BANGKOK_TZ).date()
    span = (d_to - d_from).days + 1
    is_month = d_from.day == 1 and d_to == _month_end(d_from)
    prev_to = d_from - timedelta(days=1)
    if is_month:
        prev_from = prev_to.replace(day=1)
        label = f"{THAI_MONTHS_FULL[d_from.month]} {d_from.year}"
        prev_label = f"{THAI_MONTHS[prev_from.month]} {prev_from.year}"
    else:
        prev_from = prev_to - timedelta(days=span - 1)
        label = f"{fmt_date(d_from)} – {fmt_date(d_to)}"
        prev_label = f"{fmt_date(prev_from)} – {fmt_date(prev_to)}"

    categories = get_categories(cur)
    cat_order = {c["name"]: i for i, c in enumerate(categories)}
    cat_group = {c["name"]: c["group_name"] or c["name"] for c in categories}

    def group_of(cat):
        return cat_group.get(cat, cat)

    cur.execute(
        """SELECT txn_date, txn_time, bank, direction, category, amount, fee,
                  sender_name, receiver_name, memo,
                  COALESCE(qr_trans_ref, NULLIF(printed_ref, '')), verified_bank
           FROM slip_transactions
           WHERE txn_date BETWEEN %s AND %s AND direction IN ('expense', 'unknown')
           ORDER BY txn_date, txn_time NULLS LAST, id""",
        (d_from, d_to),
    )
    txns, review = [], []
    for (t_date, t_time, bank, direction, category, amount, fee,
         sender, receiver, memo, ref, verified) in cur.fetchall():
        cat = category or UNCATEGORIZED
        t = {
            "date": t_date, "time": t_time.strftime("%H:%M") if t_time else "",
            "bank": bank or "", "direction": direction, "category": cat, "group": group_of(cat),
            "amount": float(amount or 0), "fee": float(fee or 0),
            "sender": sender or "", "receiver": receiver or "", "memo": memo or "",
            "ref": ref or "", "verified": bool(verified),
        }
        if direction == "unknown":
            t["issue"] = "ไม่แน่ใจว่าเป็นรายจ่าย"
            review.append(t)
            continue
        if not category:
            t["issue"] = "ยังไม่ระบุหมวด"
            review.append(t)
        txns.append(t)

    cur.execute(
        """SELECT COALESCE(category, %s), SUM(amount), COUNT(*)
           FROM slip_transactions
           WHERE direction = 'expense' AND txn_date BETWEEN %s AND %s
           GROUP BY 1""",
        (UNCATEGORIZED, prev_from, prev_to),
    )
    prev_by_cat = {r[0]: (float(r[1]), r[2]) for r in cur.fetchall()}

    total = sum(t["amount"] for t in txns)
    prev_total = sum(v[0] for v in prev_by_cat.values())
    prev_count = sum(v[1] for v in prev_by_cat.values())

    # --- group -> category breakdown --------------------------------------
    cat_rows = {}
    for t in txns:
        row = cat_rows.setdefault(t["category"], {"name": t["category"], "cur": 0.0, "count": 0})
        row["cur"] += t["amount"]
        row["count"] += 1
    for name in prev_by_cat:
        cat_rows.setdefault(name, {"name": name, "cur": 0.0, "count": 0})
    groups = {}
    for name, row in cat_rows.items():
        row["prev"] = prev_by_cat.get(name, (0.0, 0))[0]
        row["change"] = pct_change(row["cur"], row["prev"])
        row["share"] = row["cur"] / total * 100 if total else 0.0
        gname = group_of(name)
        g = groups.setdefault(gname, {"name": gname, "cats": [], "cur": 0.0, "prev": 0.0, "count": 0})
        g["cats"].append(row)
        g["cur"] += row["cur"]
        g["prev"] += row["prev"]
        g["count"] += row["count"]
    for g in groups.values():
        g["change"] = pct_change(g["cur"], g["prev"])
        g["share"] = g["cur"] / total * 100 if total else 0.0
        g["cats"].sort(key=lambda r: (-r["cur"], -r["prev"], cat_order.get(r["name"], 999)))
        # A category with no group is its own one-member group: render it as a single row.
        g["single"] = len(g["cats"]) == 1 and g["cats"][0]["name"] == g["name"]
    group_list = sorted(groups.values(), key=lambda g: (-g["cur"], -g["prev"]))

    # --- trend (daily up to ~2 months, monthly beyond) ----------------------
    if span <= 62:
        trend_unit = "day"
        keys = [d_from + timedelta(days=i) for i in range(span)]
        by_key = defaultdict(float)
        for t in txns:
            by_key[t["date"]] += t["amount"]
        trend = [{"key": k, "label": fmt_date(k, year=False), "short": str(k.day), "total": by_key[k]} for k in keys]
    else:
        trend_unit = "month"
        by_key = defaultdict(float)
        for t in txns:
            by_key[(t["date"].year, t["date"].month)] += t["amount"]
        trend, y, m = [], d_from.year, d_from.month
        while (y, m) <= (d_to.year, d_to.month):
            trend.append({
                "key": date(y, m, 1), "label": f"{THAI_MONTHS[m]} {y}",
                "short": f"{THAI_MONTHS[m]} {str(y)[2:]}", "total": by_key[(y, m)],
            })
            y, m = (y + 1, 1) if m == 12 else (y, m + 1)

    # --- payees & largest slips ---------------------------------------------
    payees = {}
    for t in txns:
        name = t["receiver"] or UNKNOWN_PAYEE
        p = payees.setdefault(name, {"name": name, "total": 0.0, "count": 0, "cats": defaultdict(float)})
        p["total"] += t["amount"]
        p["count"] += 1
        p["cats"][t["category"]] += t["amount"]
    top_payees = sorted(payees.values(), key=lambda p: -p["total"])[:10]
    for p in top_payees:
        p["share"] = p["total"] / total * 100 if total else 0.0
        p["main_category"] = max(p["cats"], key=p["cats"].get)
    largest = sorted(txns, key=lambda t: -t["amount"])[:10]

    # --- KPIs ---------------------------------------------------------------
    partial = d_from <= today < d_to
    days_elapsed = (today - d_from).days + 1 if partial else span
    peak = max(trend, key=lambda b: b["total"]) if trend and total else None
    kpi = {
        "total": total, "count": len(txns),
        "avg_txn": total / len(txns) if txns else 0.0,
        "avg_day": total / days_elapsed if days_elapsed else 0.0,
        "days": days_elapsed,
        "prev_total": prev_total, "prev_count": prev_count,
        "change": pct_change(total, prev_total), "diff": total - prev_total,
        "fees": sum(t["fee"] for t in txns),
        "payee_count": len(payees),
        "peak": peak,
        "review_count": len(review),
        "review_total": sum(t["amount"] for t in review),
        "uncategorized_count": sum(1 for t in review if t["direction"] == "expense"),
        "unknown_count": sum(1 for t in review if t["direction"] == "unknown"),
    }

    return {
        "from": d_from, "to": d_to, "span": span, "is_month": is_month, "partial": partial, "today": today,
        "label": label, "range_label": f"{fmt_date(d_from)} – {fmt_date(d_to)}",
        "prev_from": prev_from, "prev_to": prev_to, "prev_label": prev_label,
        "generated_at": datetime.now(BANGKOK_TZ),
        "kpi": kpi, "groups": group_list, "trend": trend, "trend_unit": trend_unit,
        "top_payees": top_payees, "largest": largest, "txns": txns, "review": review,
    }


def pie_slices(rep):
    """Group totals for the share chart: every group if they fit the palette, otherwise the
    biggest PIE_SLICES groups with the rest lumped into one grey slice."""
    groups = [g for g in rep["groups"] if g["cur"] > 0]
    shown = len(groups) if len(groups) <= len(PALETTE) else PIE_SLICES
    slices = [(g["name"], g["cur"], "#" + PALETTE[i]) for i, g in enumerate(groups[:shown])]
    rest = sum(g["cur"] for g in groups[shown:])
    if rest:
        slices.append((f"กลุ่มอื่นรวม {len(groups) - shown} กลุ่ม", rest, "#" + OTHER_COLOR))
    return slices


def txns_by_category(rep):
    """Transactions bucketed per category, in the same group/category order as the breakdown table."""
    buckets = defaultdict(list)
    for t in rep["txns"]:
        buckets[t["category"]].append(t)
    out = []
    for g in rep["groups"]:
        for c in g["cats"]:
            if buckets.get(c["name"]):
                out.append((g, c, buckets[c["name"]]))
    return out


def change_word(rep):
    k = rep["kpi"]
    if k["change"] is None:
        return f"ไม่มีข้อมูล{rep['prev_label']}ให้เทียบ"
    direction = "เพิ่มขึ้น" if k["diff"] > 0 else "ลดลง"
    return f"{direction} {abs(k['change']):.1f}% ({fmt_baht(abs(k['diff']))} บาท) เทียบ {rep['prev_label']}"


def report_filename(rep, ext):
    if rep["is_month"]:
        return f"expense_report_{rep['from']:%Y-%m}.{ext}"
    return f"expense_report_{rep['from']:%Y-%m-%d}_{rep['to']:%Y-%m-%d}.{ext}"


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

XL_FONT = "Leelawadee UI"
_thin = Side(style="thin", color=C_BORDER)
_medium = Side(style="medium", color=C_INK)
MONEY = '#,##0.00;[Red]-#,##0.00'
PCT_SIGNED = '+0.0%;-0.0%;0.0%'


def _f(size=10, bold=False, color=C_INK, italic=False):
    return Font(name=XL_FONT, size=size, bold=bold, color=color, italic=italic)


def _fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def _put(ws, row, col, value, font=None, fill=None, fmt=None, align=None, border=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = font or _f()
    if fill:
        cell.fill = fill
    if fmt:
        cell.number_format = fmt
    if align:
        cell.alignment = align
    if border:
        cell.border = border
    return cell


def _xl_title(ws, rep, title, last_col):
    ws.sheet_view.showGridLines = False
    for col in range(1, last_col + 1):
        ws.cell(row=1, column=col).fill = _fill(C_HEADER_BAND)
        ws.cell(row=2, column=col).fill = _fill(C_HEADER_BAND)
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 20
    _put(ws, 1, 1, f"{title} · {rep['label']}", _f(16, True, "FFFFFF"), _fill(C_HEADER_BAND),
         align=Alignment(vertical="center", indent=1))
    sub = f"ช่วงวันที่ {rep['range_label']}   ·   จัดทำเมื่อ {fmt_date(rep['generated_at'].date())} {rep['generated_at']:%H:%M} น."
    _put(ws, 2, 1, sub, _f(9, color="D1D5DB"), _fill(C_HEADER_BAND), align=Alignment(vertical="top", indent=1))


def _xl_section(ws, row, text, last_col):
    _put(ws, row, 1, text, _f(12, True, C_ACCENT))
    for col in range(1, last_col + 1):
        ws.cell(row=row, column=col).border = Border(bottom=Side(style="medium", color=C_ACCENT))
    ws.row_dimensions[row].height = 22
    return row + 1


def _xl_header(ws, row, headers, aligns=None):
    for i, h in enumerate(headers, start=1):
        right = aligns and aligns[i - 1] == "r"
        _put(ws, row, i, h, _f(9, True, C_MUTED), _fill("F3F4F6"),
             align=Alignment(horizontal="right" if right else "left", vertical="center", wrap_text=True),
             border=Border(bottom=_thin))
    ws.row_dimensions[row].height = 20
    return row + 1


def _xl_print_setup(ws, rep, landscape_mode=False, repeat_row=None):
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "landscape" if landscape_mode else "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.6
    ws.oddFooter.left.text = f"{REPORT_TITLE} · {rep['label']}"
    ws.oddFooter.left.size = 8
    ws.oddFooter.right.text = "หน้า &P / &N"
    ws.oddFooter.right.size = 8
    if repeat_row:
        ws.print_title_rows = f"{repeat_row}:{repeat_row}"


def _change_font(change, size=10, bold=False):
    if change is None:
        return _f(size, bold, C_FAINT)
    return _f(size, bold, C_UP if change > 0 else C_DOWN)


def _xl_summary_sheet(wb, rep, chart_ws):
    ws = wb.active
    ws.title = "สรุป"
    k = rep["kpi"]
    widths = {"A": 40, "B": 11, "C": 17, "D": 17, "E": 12, "F": 11}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    last = 6
    _xl_title(ws, rep, REPORT_TITLE, last)

    # --- KPI block -----------------------------------------------------------
    row = _xl_section(ws, 4, "ภาพรวม", last)
    kpis = [
        ("รายจ่ายรวม", k["total"], MONEY, "บาท"),
        ("จำนวนรายการ", k["count"], "#,##0", f"สลิป · ผู้รับ {k['payee_count']} ราย"),
        ("เฉลี่ยต่อรายการ", k["avg_txn"], MONEY, "บาท"),
        ("เฉลี่ยต่อวัน", k["avg_day"], MONEY, f"บาท (คิดจาก {k['days']} วัน)"),
        (f"ช่วงก่อนหน้า ({rep['prev_label']})", k["prev_total"], MONEY, f"บาท · {k['prev_count']} รายการ"),
        ("เปลี่ยนแปลง", (k["change"] / 100) if k["change"] is not None else "–", PCT_SIGNED,
         f"{'+' if k['diff'] > 0 else ''}{fmt_baht(k['diff'])} บาท"),
    ]
    if k["peak"]:
        unit = "วัน" if rep["trend_unit"] == "day" else "เดือน"
        kpis.append((f"{unit}ที่จ่ายมากที่สุด", k["peak"]["total"], MONEY, k["peak"]["label"]))
    if k["fees"]:
        kpis.append(("ค่าธรรมเนียมโอนรวม", k["fees"], MONEY, "บาท"))
    for label, value, fmt, note in kpis:
        is_total = label == "รายจ่ายรวม"
        _put(ws, row, 1, label, _f(10, color=C_MUTED), align=Alignment(indent=1, vertical="center"))
        if label == "เปลี่ยนแปลง":
            vfont = _change_font(k["change"], 11, True)
        else:
            vfont = _f(14 if is_total else 11, True)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
        _put(ws, row, 2, value, vfont, fmt=fmt, align=Alignment(horizontal="right", vertical="center"))
        _put(ws, row, 4, note, _f(9, color=C_MUTED), align=Alignment(indent=1, vertical="center"))
        for col in range(1, last + 1):
            ws.cell(row=row, column=col).border = Border(bottom=_thin)
        ws.row_dimensions[row].height = 24 if is_total else 19
        row += 1
    if rep["partial"]:
        _put(ws, row, 1, f"หมายเหตุ: ช่วงนี้ยังไม่จบ ข้อมูลถึง {fmt_date(rep['today'])} — การเทียบกับช่วงก่อนหน้าเป็นยอดไม่เต็มช่วง",
             _f(9, color=C_WARN, italic=True), align=Alignment(indent=1))
        row += 1
    if k["review_count"]:
        parts = []
        if k["uncategorized_count"]:
            parts.append(f"ยังไม่ระบุหมวด {k['uncategorized_count']} รายการ")
        if k["unknown_count"]:
            parts.append(f"ไม่แน่ใจว่าเป็นรายจ่าย {k['unknown_count']} รายการ (ไม่รวมในยอด)")
        for col in range(1, last + 1):
            ws.cell(row=row, column=col).fill = _fill(C_WARN_SOFT)
        _put(ws, row, 1, "ต้องตรวจ: " + " · ".join(parts) + " — ดูชีต 'ต้องตรวจ'",
             _f(10, True, C_WARN), _fill(C_WARN_SOFT), align=Alignment(indent=1, vertical="center"))
        ws.row_dimensions[row].height = 22
        row += 1

    # --- group / category table ----------------------------------------------
    row = _xl_section(ws, row + 1, "แยกตามกลุ่มและหมวด", last)
    row = _xl_header(ws, row, ["กลุ่ม / หมวด", "จำนวน", rep["label"] if rep["is_month"] else "ช่วงนี้",
                               rep["prev_label"] if rep["is_month"] else "ช่วงก่อน", "เปลี่ยนแปลง", "สัดส่วน"],
                     "lrrrrr")
    table_top = row

    def line(r, name, d, bold=False, indent=1, fill=None):
        _put(ws, r, 1, name, _f(10, bold), fill, align=Alignment(indent=indent))
        _put(ws, r, 2, d["count"], _f(10, bold), fill, "#,##0", Alignment(horizontal="right"))
        _put(ws, r, 3, d["cur"], _f(10, bold), fill, MONEY)
        _put(ws, r, 4, d["prev"], _f(10, bold, C_MUTED), fill, MONEY)
        _put(ws, r, 5, d["change"] / 100 if d["change"] is not None else "–",
             _change_font(d["change"], 10, bold), fill, PCT_SIGNED, Alignment(horizontal="right"))
        _put(ws, r, 6, d["share"] / 100, _f(10, bold), fill, "0.0%")
        for col in range(1, last + 1):
            ws.cell(row=r, column=col).border = Border(bottom=_thin)

    for g in rep["groups"]:
        if g["single"]:
            line(row, g["name"], g, bold=True)
            row += 1
            continue
        line(row, g["name"], g, bold=True, fill=_fill(C_SURFACE_ALT))
        row += 1
        for c in g["cats"]:
            line(row, c["name"], c, indent=3)
            ws.row_dimensions[row].outlineLevel = 1
            row += 1
    totals = {"count": k["count"], "cur": k["total"], "prev": k["prev_total"], "change": k["change"],
              "share": 100.0 if k["total"] else 0.0}
    line(row, "รวมทั้งหมด", totals, bold=True, fill=_fill(C_ACCENT_SOFT))
    for col in range(1, last + 1):
        ws.cell(row=row, column=col).border = Border(top=_medium, bottom=_medium)
    ws.row_dimensions[row].height = 20
    row += 2
    ws.sheet_properties.outlinePr.summaryBelow = False

    # --- top payees ------------------------------------------------------------
    row = _xl_section(ws, row, "ผู้รับเงินสูงสุด 10 อันดับ", last)
    row = _xl_header(ws, row, ["ผู้รับ", "จำนวน", "ยอดรวม", "หมวดหลัก", "", "สัดส่วน"], "lrrlrr")
    for i, p in enumerate(rep["top_payees"], start=1):
        _put(ws, row, 1, f"{i}.  {p['name']}", align=Alignment(indent=1))
        _put(ws, row, 2, p["count"], fmt="#,##0", align=Alignment(horizontal="right"))
        _put(ws, row, 3, p["total"], _f(10, True), fmt=MONEY)
        ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=5)
        _put(ws, row, 4, p["main_category"], _f(9, color=C_MUTED), align=Alignment(indent=1))
        _put(ws, row, 6, p["share"] / 100, fmt="0.0%")
        for col in range(1, last + 1):
            ws.cell(row=row, column=col).border = Border(bottom=_thin)
        row += 1

    # --- largest slips -----------------------------------------------------------
    row = _xl_section(ws, row + 1, "รายการใหญ่ที่สุด 10 อันดับ", last)
    row = _xl_header(ws, row, ["ผู้รับ / รายละเอียด", "วันที่", "ยอดเงิน", "หมวด", "", ""], "lrrlll")
    for t in rep["largest"]:
        desc = t["receiver"] or UNKNOWN_PAYEE
        if t["memo"]:
            desc += f" — {t['memo']}"
        _put(ws, row, 1, desc, align=Alignment(indent=1))
        _put(ws, row, 2, t["date"], fmt="dd/mm/yyyy", align=Alignment(horizontal="right"))
        _put(ws, row, 3, t["amount"], _f(10, True), fmt=MONEY)
        ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=6)
        _put(ws, row, 4, t["category"], _f(9, color=C_MUTED), align=Alignment(indent=1))
        for col in range(1, last + 1):
            ws.cell(row=row, column=col).border = Border(bottom=_thin)
        row += 1

    # --- charts (data lives on the chart-data sheet) ------------------------------
    ws.column_dimensions["G"].width = 3
    n_slices = len(pie_slices(rep))
    if n_slices:
        pie = PieChart()
        pie.title = "สัดส่วนรายจ่ายตามกลุ่ม"
        pie.add_data(Reference(chart_ws, min_col=2, min_row=1, max_row=n_slices + 1), titles_from_data=True)
        pie.set_categories(Reference(chart_ws, min_col=1, min_row=2, max_row=n_slices + 1))
        for i, (_, _, color) in enumerate(pie_slices(rep)):
            pt = DataPoint(idx=i)
            pt.graphicalProperties.solidFill = color.lstrip("#")
            pt.graphicalProperties.line.solidFill = "FFFFFF"
            pie.series[0].dPt.append(pt)
        # Percentages live in the legend labels (see chart-data sheet) — on-slice labels collide.
        pie.title.overlay = False
        pie.legend.position = "r"
        pie.legend.overlay = False
        pie.width, pie.height = 17, 9
        ws.add_chart(pie, "H4")
    if rep["trend"]:
        n = len(rep["trend"])
        bar = BarChart()
        bar.type = "col"
        bar.title = "รายจ่ายรายวัน" if rep["trend_unit"] == "day" else "รายจ่ายรายเดือน"
        bar.add_data(Reference(chart_ws, min_col=5, min_row=1, max_row=n + 1), titles_from_data=True)
        bar.set_categories(Reference(chart_ws, min_col=4, min_row=2, max_row=n + 1))
        bar.series[0].graphicalProperties.solidFill = C_ACCENT
        bar.series[0].graphicalProperties.line.noFill = True
        bar.title.overlay = False
        bar.legend = None
        bar.gapWidth = 40
        bar.y_axis.numFmt = "#,##0"
        # openpyxl >= 3.1 writes axes as deleted unless told otherwise.
        bar.x_axis.delete = False
        bar.y_axis.delete = False
        bar.width, bar.height = 17, 8
        ws.add_chart(bar, "H23")

    ws.freeze_panes = "A3"
    _xl_print_setup(ws, rep)
    # Charts sit beside the table; printing them would shrink the table to fit the page width.
    ws.print_area = f"A1:F{row}"


def _xl_chart_data_sheet(wb, rep):
    ws = wb.create_sheet("ข้อมูลกราฟ")
    _put(ws, 1, 1, "กลุ่ม", _f(9, True))
    _put(ws, 1, 2, "ยอด (บาท)", _f(9, True))
    slices = pie_slices(rep)
    total = sum(s[1] for s in slices)
    for i, (name, value, _) in enumerate(slices, start=2):
        _put(ws, i, 1, f"{name}  {value / total * 100:.0f}%")
        _put(ws, i, 2, value, fmt=MONEY)
    _put(ws, 1, 4, "วันที่" if rep["trend_unit"] == "day" else "เดือน", _f(9, True))
    _put(ws, 1, 5, "รายจ่าย (บาท)", _f(9, True))
    for i, b in enumerate(rep["trend"], start=2):
        _put(ws, i, 4, b["short"])
        _put(ws, i, 5, b["total"], fmt=MONEY)
    for col, w in {"A": 34, "B": 14, "D": 14, "E": 16}.items():
        ws.column_dimensions[col].width = w
    ws.sheet_state = "hidden"
    return ws


def _xl_by_category_sheet(wb, rep):
    ws = wb.create_sheet("ตามหมวด")
    widths = {"A": 13, "B": 8, "C": 34, "D": 44, "E": 12, "F": 15}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    last = 6
    _xl_title(ws, rep, "รายละเอียดตามหมวด", last)
    row = _xl_header(ws, 4, ["วันที่", "เวลา", "ผู้รับ", "รายละเอียด", "ธนาคาร", "ยอดเงิน"], "lllllr")
    ws.freeze_panes = "A5"
    for g, c, items in txns_by_category(rep):
        title = c["name"] if c["name"] == g["name"] else f"{c['name']}   ({g['name']})"
        for col in range(1, last + 1):
            ws.cell(row=row, column=col).fill = _fill(C_ACCENT_SOFT)
        _put(ws, row, 1, title, _f(11, True, C_ACCENT), _fill(C_ACCENT_SOFT), align=Alignment(vertical="center"))
        _put(ws, row, 5, f"{len(items)} รายการ", _f(9, color=C_MUTED), _fill(C_ACCENT_SOFT),
             align=Alignment(horizontal="right", vertical="center"))
        _put(ws, row, 6, c["cur"], _f(11, True, C_ACCENT), _fill(C_ACCENT_SOFT), MONEY, Alignment(vertical="center"))
        ws.row_dimensions[row].height = 22
        row += 1
        for t in items:
            _put(ws, row, 1, t["date"], fmt="dd/mm/yyyy", align=Alignment(horizontal="left", vertical="top"))
            _put(ws, row, 2, t["time"], _f(10, color=C_MUTED), align=Alignment(vertical="top"))
            _put(ws, row, 3, t["receiver"] or UNKNOWN_PAYEE, align=Alignment(wrap_text=True, vertical="top"))
            _put(ws, row, 4, t["memo"], _f(10, color=C_MUTED), align=Alignment(wrap_text=True, vertical="top"))
            _put(ws, row, 5, t["bank"], _f(9, color=C_MUTED), align=Alignment(vertical="top"))
            _put(ws, row, 6, t["amount"], fmt=MONEY, align=Alignment(vertical="top"))
            for col in range(1, last + 1):
                ws.cell(row=row, column=col).border = Border(bottom=_thin)
            ws.row_dimensions[row].outlineLevel = 1
            row += 1
        row += 1
    for col in range(1, last + 1):
        ws.cell(row=row, column=col).border = Border(top=_medium, bottom=_medium)
        ws.cell(row=row, column=col).fill = _fill(C_ACCENT_SOFT)
    _put(ws, row, 1, "รวมทั้งหมด", _f(11, True), _fill(C_ACCENT_SOFT))
    _put(ws, row, 5, f"{rep['kpi']['count']} รายการ", _f(9, color=C_MUTED), _fill(C_ACCENT_SOFT),
         align=Alignment(horizontal="right"))
    _put(ws, row, 6, rep["kpi"]["total"], _f(11, True), _fill(C_ACCENT_SOFT), MONEY)
    ws.sheet_properties.outlinePr.summaryBelow = False
    _xl_print_setup(ws, rep, repeat_row=4)


TXN_COLUMNS = [
    ("วันที่", 12), ("เวลา", 7), ("ธนาคาร", 14), ("กลุ่ม", 22), ("หมวด", 26), ("ผู้รับ", 30),
    ("รายละเอียด", 36), ("ผู้โอน", 22), ("ยอดเงิน", 14), ("ค่าธรรมเนียม", 12),
    ("เลขอ้างอิง", 26), ("ยืนยันกับธนาคาร", 12),
]


def _txn_row(t):
    return [t["date"], t["time"], t["bank"], t["group"], t["category"], t["receiver"], t["memo"],
            t["sender"], t["amount"], t["fee"], t["ref"], "ใช่" if t["verified"] else "ไม่"]


def _xl_flat_table(ws, rep, rows, table_name, top, extra_cols=()):
    columns = TXN_COLUMNS + list(extra_cols)
    for i, (name, width) in enumerate(columns, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    _xl_header(ws, top, [c[0] for c in columns], ["r" if c[0] in ("ยอดเงิน", "ค่าธรรมเนียม") else "l" for c in columns])
    r = top
    for r, values in enumerate(rows, start=top + 1):
        for col, v in enumerate(values, start=1):
            name = columns[col - 1][0]
            fmt = "dd/mm/yyyy" if name == "วันที่" else MONEY if name in ("ยอดเงิน", "ค่าธรรมเนียม") else None
            _put(ws, r, col, v, fmt=fmt, align=Alignment(vertical="top", horizontal="left" if name == "วันที่" else None))
    if rows:
        ref = f"A{top}:{get_column_letter(len(columns))}{r}"
        tbl = XLTable(displayName=table_name, ref=ref)
        tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
        ws.add_table(tbl)
    return r, columns


def _xl_all_txns_sheet(wb, rep):
    ws = wb.create_sheet("รายการทั้งหมด")
    n_cols = len(TXN_COLUMNS)
    _xl_title(ws, rep, "รายการรายจ่ายทั้งหมด", n_cols)
    last_row, columns = _xl_flat_table(ws, rep, [_txn_row(t) for t in rep["txns"]], "tbl_txns", 4)
    if rep["txns"]:
        # SUBTOTAL(109, ...) follows the AutoFilter, so the total tracks whatever is filtered.
        names = [c[0] for c in columns]
        amt_idx, fee_idx = names.index("ยอดเงิน") + 1, names.index("ค่าธรรมเนียม") + 1
        amt_col, fee_col = get_column_letter(amt_idx), get_column_letter(fee_idx)
        tr = last_row + 1
        _put(ws, tr, 1, "รวม (ตามตัวกรอง)", _f(10, True))
        _put(ws, tr, amt_idx - 1, f"=SUBTOTAL(103,{amt_col}5:{amt_col}{last_row})", _f(10, True), fmt='#,##0" รายการ"')
        _put(ws, tr, amt_idx, f"=SUBTOTAL(109,{amt_col}5:{amt_col}{last_row})", _f(10, True), fmt=MONEY)
        _put(ws, tr, fee_idx, f"=SUBTOTAL(109,{fee_col}5:{fee_col}{last_row})", _f(10, True), fmt=MONEY)
        for col in range(1, n_cols + 1):
            ws.cell(row=tr, column=col).border = Border(top=_medium)
            ws.cell(row=tr, column=col).fill = _fill(C_ACCENT_SOFT)
    ws.freeze_panes = "B5"
    _xl_print_setup(ws, rep, landscape_mode=True, repeat_row=4)


def _xl_review_sheet(wb, rep):
    if not rep["review"]:
        return
    ws = wb.create_sheet("ต้องตรวจ")
    ws.sheet_properties.tabColor = C_WARN
    columns = TXN_COLUMNS + [("ปัญหา", 24)]
    _xl_title(ws, rep, "รายการที่ต้องตรวจ", len(columns))
    _put(ws, 3, 1, "รายการที่ 'ไม่แน่ใจว่าเป็นรายจ่าย' ไม่ได้รวมในยอดรายงาน — แก้ในหน้าเว็บแล้ว export ใหม่",
         _f(9, color=C_WARN, italic=True))
    _xl_flat_table(ws, rep, [_txn_row(t) + [t["issue"]] for t in rep["review"]], "tbl_review", 4,
                   extra_cols=[("ปัญหา", 24)])
    ws.freeze_panes = "B5"
    _xl_print_setup(ws, rep, landscape_mode=True, repeat_row=4)


def build_report_xlsx(rep, output):
    """Render a fetch_period_report() dict to .xlsx. `output` is a path or a file-like object."""
    wb = Workbook()
    chart_ws = _xl_chart_data_sheet(wb, rep)
    _xl_summary_sheet(wb, rep, chart_ws)
    _xl_by_category_sheet(wb, rep)
    _xl_all_txns_sheet(wb, rep)
    _xl_review_sheet(wb, rep)
    wb.move_sheet(chart_ws, offset=len(wb.sheetnames))
    wb.active = 0
    wb.properties.title = f"{REPORT_TITLE} {rep['label']}"
    wb.save(output)


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------

# reportlab breaks lines only at spaces (or anywhere, in CJK mode — which splits Thai
# vowel/tone marks off their consonant). Thai has no spaces between words, so long
# cells are pre-wrapped here at grapheme-cluster boundaries, preferring real spaces.
_THAI_NON_STARTERS = {chr(c) for c in [0x0E30, 0x0E31, 0x0E32, 0x0E33, *range(0x0E34, 0x0E3B), 0x0E45, *range(0x0E47, 0x0E4F)]}
_THAI_LEADING_VOWELS = {chr(c) for c in range(0x0E40, 0x0E45)}


def _can_break_before(text, i):
    return i > 0 and text[i] not in _THAI_NON_STARTERS and text[i - 1] not in _THAI_LEADING_VOWELS


def wrap_text(text, font, size, width, max_lines=None):
    lines = []
    for para in str(text).splitlines() or [""]:
        rest = para.strip()
        while rest:
            if stringWidth(rest, font, size) <= width:
                lines.append(rest)
                break
            cut, last_space, last_break = 0, None, None
            for i in range(1, len(rest) + 1):
                if stringWidth(rest[:i], font, size) > width:
                    break
                cut = i
                if i < len(rest) and rest[i] == " ":
                    last_space = i
                if i < len(rest) and _can_break_before(rest, i):
                    last_break = i
            # Prefer a space in the back two-thirds of the line; otherwise a cluster boundary.
            split = last_space if last_space and last_space > cut / 3 else (last_break or cut or 1)
            lines.append(rest[:split].rstrip())
            rest = rest[split:].lstrip()
    if max_lines and len(lines) > max_lines:
        lines = lines[:max_lines]
        lines[-1] = lines[-1].rstrip() + "…"
    return lines


_styles = getSampleStyleSheet()
ST = {
    "body": _ps("rb", parent=_styles["BodyText"], fontName=FONT, fontSize=9, leading=12.5, textColor=_hex(C_INK)),
    "muted": _ps("rm", fontName=FONT, fontSize=8, leading=11, textColor=_hex(C_MUTED)),
    "h2": _ps("rh2", fontName=FONT_BOLD, fontSize=13, leading=17, textColor=_hex(C_INK), spaceBefore=14, spaceAfter=6),
    "h3": _ps("rh3", fontName=FONT_BOLD, fontSize=10.5, leading=14, textColor=_hex(C_ACCENT)),
    "kpi_label": _ps("kl", fontName=FONT, fontSize=8, leading=10, textColor=_hex(C_MUTED)),
    "kpi_value": _ps("kv", fontName=FONT_BOLD, fontSize=15, leading=19, textColor=_hex(C_INK)),
    "kpi_sub": _ps("ks", fontName=FONT, fontSize=7.5, leading=10, textColor=_hex(C_MUTED)),
    "cell": _ps("c", fontName=FONT, fontSize=8.5, leading=11, textColor=_hex(C_INK)),
    "cell_muted": _ps("cm", fontName=FONT, fontSize=8, leading=10.5, textColor=_hex(C_MUTED)),
    "cell_r": _ps("cr", fontName=FONT, fontSize=8.5, leading=11, textColor=_hex(C_INK), alignment=TA_RIGHT),
    "head": _ps("hd", fontName=FONT_BOLD, fontSize=8, leading=10, textColor=_hex(C_MUTED)),
    "head_r": _ps("hdr", fontName=FONT_BOLD, fontSize=8, leading=10, textColor=_hex(C_MUTED), alignment=TA_RIGHT),
}

PAGE_W, PAGE_H = A4
MARGIN = 1.5 * cm
CONTENT_W = PAGE_W - 2 * MARGIN


def P(text, style="cell", bold=False, color=None):
    s = escape(str(text))
    if bold:
        s = f"<b>{s}</b>"
    if color:
        s = f'<font color="#{color}">{s}</font>'
    return Paragraph(s, ST[style])


def PW(text, width, style="cell", max_lines=None, bold=False):
    """Paragraph pre-wrapped for Thai to fit `width` points (minus cell padding)."""
    st = ST[style]
    lines = wrap_text(text, FONT_BOLD if bold else st.fontName, st.fontSize, width - 8, max_lines)
    markup = "<br/>".join(escape(line) for line in lines)
    return Paragraph(f"<b>{markup}</b>" if bold else markup, st)


def _change_markup(change, bold=False):
    if change is None:
        return P("–", "cell_r", color=C_FAINT)
    return P(fmt_pct(change, signed=True), "cell_r", bold=bold, color=C_UP if change > 0 else C_DOWN)


def _base_table_style(header_rows=1):
    return [
        ("FONTNAME", (0, 0), (-1, -1), FONT),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("BACKGROUND", (0, 0), (-1, header_rows - 1), _hex("F3F4F6")),
        ("LINEBELOW", (0, header_rows - 1), (-1, header_rows - 1), 0.8, _hex(C_FAINT)),
        ("LINEBELOW", (0, header_rows), (-1, -1), 0.4, _hex(C_BORDER)),
    ]


def _share_bar(share, width=46, height=6, color=C_ACCENT):
    d = Drawing(width, height)
    d.add(Rect(0, 0, width, height, fillColor=_hex("EEF0F3"), strokeColor=None))
    d.add(Rect(0, 0, max(width * min(share, 100) / 100, 0.5), height, fillColor=_hex(color), strokeColor=None))
    return d


def _share_cell(share, bold=False, color=C_ACCENT):
    """Inline bar + percentage, for a table's share column."""
    t = Table([[_share_bar(share, 40, 5, color), P(f"{share:.1f}%", "cell_r", bold=bold)]], colWidths=[44, 37])
    t.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return t


def _nice_max(v):
    if v <= 0:
        return 1, 1
    exp = 10 ** (len(str(int(v))) - 1)
    for mult in (1, 2, 2.5, 5, 10):
        step = mult * exp / 4
        if step * 4 >= v:
            return step * 4, step
    return v, v / 4


class _NumberedCanvas(rl_canvas.Canvas):
    """Defers page output so the footer can print 'page X / Y'."""

    def __init__(self, *args, footer_left="", **kwargs):
        super().__init__(*args, **kwargs)
        self._saved = []
        self._footer_left = footer_left

    def showPage(self):
        self._saved.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._saved)
        for state in self._saved:
            self.__dict__.update(state)
            self._draw_footer(total)
            super().showPage()
        super().save()

    def _draw_footer(self, total):
        y = MARGIN - 0.75 * cm
        self.setStrokeColor(_hex(C_BORDER))
        self.setLineWidth(0.5)
        self.line(MARGIN, y + 11, PAGE_W - MARGIN, y + 11)
        self.setFont(FONT, 7.5)
        self.setFillColor(_hex(C_MUTED))
        self.drawString(MARGIN, y, self._footer_left, shaping=SHAPING)
        self.drawRightString(PAGE_W - MARGIN, y, f"หน้า {self._pageNumber} / {total}", shaping=SHAPING)


def _pdf_header(rep):
    gen = rep["generated_at"]
    title = Paragraph(
        f'<font size="9" color="#9CA3AF">{REPORT_TITLE.upper()}</font><br/>'
        f'<font name="{FONT_BOLD}" size="22" color="#FFFFFF">{escape(rep["label"])}</font>',
        _ps("hdr_t", fontName=FONT, fontSize=22, leading=28, textColor=colors.white),
    )
    meta = Paragraph(
        f"ช่วงวันที่ {escape(rep['range_label'])}<br/>"
        f"เทียบกับ {escape(rep['prev_label'])}<br/>"
        f"จัดทำเมื่อ {fmt_date(gen.date())} {gen:%H:%M} น.",
        _ps("hdr_m", fontName=FONT, fontSize=8, leading=12, textColor=_hex("D1D5DB"), alignment=TA_RIGHT),
    )
    t = Table([[title, meta]], colWidths=[CONTENT_W * 0.6, CONTENT_W * 0.4])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), _hex(C_HEADER_BAND)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 14),
        ("RIGHTPADDING", (0, 0), (-1, -1), 14),
        ("TOPPADDING", (0, 0), (-1, -1), 12),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 14),
        ("ROUNDEDCORNERS", [6, 6, 6, 6]),
    ]))
    return t


def _pdf_kpis(rep):
    k = rep["kpi"]
    if k["change"] is None:
        change = f'<font color="#{C_FAINT}">ไม่มีข้อมูลช่วงก่อนให้เทียบ</font>'
    else:
        color = C_UP if k["diff"] > 0 else C_DOWN
        change = f'<font color="#{color}"><b>{fmt_pct(k["change"], signed=True)}</b></font> เทียบ {escape(rep["prev_label"])}'
    peak = ""
    if k["peak"]:
        peak = f"สูงสุด {escape(k['peak']['label'])} · {k['peak']['total']:,.0f}"
    cards = [
        ("รายจ่ายรวม (บาท)", fmt_baht(k["total"]), change),
        ("จำนวนรายการ", f"{k['count']:,}", f"ผู้รับ {k['payee_count']} ราย · เฉลี่ย {k['avg_txn']:,.0f}/รายการ"),
        (f"{escape(rep['prev_label'])} (บาท)" if rep["is_month"] else "ช่วงก่อนหน้า (บาท)",
         fmt_baht(k["prev_total"]), f"{k['prev_count']} รายการ"),
        ("เฉลี่ยต่อวัน (บาท)", fmt_baht(k["avg_day"]), peak or f"คิดจาก {k['days']} วัน"),
    ]
    gap = 6
    card_w = (CONTENT_W - gap * 3) / 4
    cells, widths = [], []
    for i, (label, value, sub) in enumerate(cards):
        cells.append([
            Paragraph(label, ST["kpi_label"]),
            Paragraph(value, ST["kpi_value"]),
            Paragraph(sub, ST["kpi_sub"]),
        ])
        widths.append(card_w)
        if i < len(cards) - 1:
            cells.append("")
            widths.append(gap)
    t = Table([cells], colWidths=widths)
    style = [
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
    ]
    for i in range(0, len(cells), 2):
        style += [
            ("BOX", (i, 0), (i, 0), 0.6, _hex(C_BORDER)),
            ("BACKGROUND", (i, 0), (i, 0), colors.white),
        ]
    style.append(("LINEABOVE", (0, 0), (0, 0), 2.5, _hex(C_ACCENT)))
    t.setStyle(TableStyle(style))
    return t


def _callout(text, fg=C_WARN, bg=C_WARN_SOFT):
    t = Table([[Paragraph(text, _ps("co", fontName=FONT, fontSize=8.5, leading=12, textColor=_hex(fg)))]],
              colWidths=[CONTENT_W])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), _hex(bg)),
        ("LINEBEFORE", (0, 0), (0, -1), 3, _hex(fg)),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    return t


def _pdf_highlights(rep):
    """Plain-language takeaways, so the reader gets the story before the tables."""
    k = rep["kpi"]
    out = [f"รายจ่ายรวม <b>{fmt_baht(k['total'])}</b> บาท จาก {k['count']} รายการ — {escape(change_word(rep))}"]
    top = [g for g in rep["groups"] if g["cur"] > 0][:3]
    if top:
        parts = ", ".join(f"{escape(g['name'])} {g['share']:.0f}%" for g in top)
        out.append(f"กลุ่มที่ใช้มากที่สุด: {parts}")
    movers = [g for g in rep["groups"] if g["prev"] or g["cur"]]
    up = max(movers, key=lambda g: g["cur"] - g["prev"], default=None)
    down = min(movers, key=lambda g: g["cur"] - g["prev"], default=None)
    if up and up["cur"] - up["prev"] > 0:
        out.append(f"เพิ่มขึ้นมากสุด: {escape(up['name'])} <font color='#{C_UP}'>+{fmt_baht(up['cur'] - up['prev'])}</font> บาท")
    if down and down["cur"] - down["prev"] < 0:
        out.append(f"ลดลงมากสุด: {escape(down['name'])} <font color='#{C_DOWN}'>-{fmt_baht(down['prev'] - down['cur'])}</font> บาท")
    if rep["top_payees"]:
        p = rep["top_payees"][0]
        out.append(f"ผู้รับเงินมากที่สุด: {escape(p['name'])} {fmt_baht(p['total'])} บาท ({p['count']} รายการ)")
    if rep["largest"]:
        t = rep["largest"][0]
        out.append(f"รายการใหญ่สุด: {fmt_baht(t['amount'])} บาท — {escape(t['receiver'] or UNKNOWN_PAYEE)} ({fmt_date(t['date'])})")
    style = _ps("hl", fontName=FONT, fontSize=9, leading=13.5, textColor=_hex(C_INK), leftIndent=10, bulletIndent=0)
    return [Paragraph(s, style, bulletText="•") for s in out]


def _pdf_pie(rep):
    slices = pie_slices(rep)
    if not slices:
        return None
    total = sum(s[1] for s in slices)
    size = 120
    d = Drawing(size + 16, size + 10)
    pie = Pie()
    pie.x, pie.y = 8, 5
    pie.width = pie.height = size
    pie.data = [s[1] for s in slices]
    pie.simpleLabels = 1
    pie.labels = None
    pie.startAngle = 90
    pie.direction = "clockwise"
    for i, s in enumerate(slices):
        pie.slices[i].fillColor = colors.HexColor(s[2])
        pie.slices[i].strokeColor = colors.white
        pie.slices[i].strokeWidth = 1.2
    d.add(pie)
    cx, cy = pie.x + size / 2, pie.y + size / 2
    d.add(Circle(cx, cy, size * 0.31, fillColor=colors.white, strokeColor=None))
    d.add(String(cx, cy + 1, f"{total:,.0f}", fontName=FONT_BOLD, fontSize=10.5, fillColor=_hex(C_INK), textAnchor="middle"))
    d.add(String(cx, cy - 11, "บาท", fontName=FONT, fontSize=7.5, fillColor=_hex(C_MUTED), textAnchor="middle"))

    # Legend as a table of Paragraphs (Drawing strings aren't HarfBuzz-shaped).
    legend_w = CONTENT_W - d.width - 10
    widths = [16, legend_w - 16 - 80 - 50, 80, 50]
    rows = []
    for name, value, color in slices:
        swatch = Drawing(9, 9)
        swatch.add(Rect(0, 0, 9, 9, fillColor=colors.HexColor(color), strokeColor=None))
        rows.append([swatch, PW(name, widths[1], max_lines=1), P(fmt_baht(value), "cell_r"),
                     P(f"{value / total * 100:.1f}%", "cell_r", bold=True, color=C_MUTED)])
    legend = Table(rows, colWidths=widths)
    legend.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
        ("LINEBELOW", (0, 0), (-1, -2), 0.4, _hex(C_BORDER)),
    ]))
    t = Table([[d, legend]], colWidths=[d.width + 10, legend_w])
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    return t


def _pdf_trend(rep):
    trend = rep["trend"]
    if not trend or not rep["kpi"]["total"]:
        return None
    height = 150
    d = Drawing(CONTENT_W, height)
    chart = VerticalBarChart()
    chart.x, chart.y = 42, 22
    chart.width, chart.height = CONTENT_W - 50, height - 34
    chart.data = [[b["total"] for b in trend]]
    vmax, step = _nice_max(max(b["total"] for b in trend))
    chart.valueAxis.valueMin = 0
    chart.valueAxis.valueMax = vmax
    chart.valueAxis.valueStep = step
    chart.valueAxis.labels.fontName = FONT
    chart.valueAxis.labels.fontSize = 7
    chart.valueAxis.labels.fillColor = _hex(C_MUTED)
    chart.valueAxis.labelTextFormat = lambda v: f"{v:,.0f}"
    chart.valueAxis.strokeColor = None
    chart.valueAxis.visibleGrid = 1
    chart.valueAxis.gridStrokeColor = _hex(C_BORDER)
    chart.valueAxis.gridStrokeWidth = 0.4
    many = len(trend) > 16
    chart.categoryAxis.categoryNames = [
        b["short"] if not many or i % 5 == 0 or i == len(trend) - 1 else "" for i, b in enumerate(trend)
    ]
    chart.categoryAxis.labels.fontName = FONT
    chart.categoryAxis.labels.fontSize = 7
    chart.categoryAxis.labels.fillColor = _hex(C_MUTED)
    chart.categoryAxis.strokeColor = _hex(C_FAINT)
    chart.categoryAxis.tickDown = 0
    chart.bars[0].fillColor = _hex(C_ACCENT)
    chart.bars[0].strokeColor = None
    chart.barSpacing = 0
    chart.groupSpacing = 3 if many else 8
    d.add(chart)
    # dashed average line
    avg = rep["kpi"]["total"] / len(trend)
    y = chart.y + avg / vmax * chart.height
    d.add(Line(chart.x, y, chart.x + chart.width, y, strokeColor=_hex(C_UP), strokeWidth=0.8, strokeDashArray=[3, 2]))
    unit = "วัน" if rep["trend_unit"] == "day" else "เดือน"
    d.add(String(chart.x + chart.width, y + 3, f"เฉลี่ย {avg:,.0f}/{unit}", fontName=FONT, fontSize=7,
                 fillColor=_hex(C_UP), textAnchor="end"))
    return d


def _pdf_breakdown(rep):
    k = rep["kpi"]
    widths = [CONTENT_W - 330, 40, 75, 75, 55, 85]
    head_cur = "ช่วงนี้" if not rep["is_month"] else rep["label"].split()[0]
    rows = [[P("กลุ่ม / หมวด", "head"), P("รายการ", "head_r"), P(head_cur, "head_r"),
             P(rep["prev_label"] if rep["is_month"] else "ช่วงก่อน", "head_r"), P("เปลี่ยน", "head_r"), P("สัดส่วน", "head")]]
    style = _base_table_style()

    def add(name, d, bold=False, indent=False, bg=None):
        r = len(rows)
        label_w = widths[0] - (14 if indent else 0)
        name_p = PW(name, label_w, "cell", bold=bold)
        share = _share_cell(d["share"], bold, C_ACCENT if bold else C_FAINT)
        rows.append([
            name_p, P(f"{d['count']:,}", "cell_r", bold=bold), P(fmt_baht(d["cur"]), "cell_r", bold=bold),
            P(fmt_baht(d["prev"]), "cell_r", color=C_MUTED), _change_markup(d["change"], bold), share,
        ])
        if indent:
            style.append(("LEFTPADDING", (0, r), (0, r), 18))
        if bg:
            style.append(("BACKGROUND", (0, r), (-1, r), _hex(bg)))

    for g in rep["groups"]:
        if g["single"]:
            add(g["name"], g, bold=True)
            continue
        add(g["name"], g, bold=True, bg=C_SURFACE_ALT)
        for c in g["cats"]:
            add(c["name"], c, indent=True)
    add("รวมทั้งหมด", {"count": k["count"], "cur": k["total"], "prev": k["prev_total"],
                       "change": k["change"], "share": 100.0 if k["total"] else 0.0}, bold=True, bg=C_ACCENT_SOFT)
    style += [("LINEABOVE", (0, -1), (-1, -1), 1, _hex(C_INK)), ("VALIGN", (5, 0), (5, -1), "MIDDLE")]
    t = Table(rows, colWidths=widths, repeatRows=1)
    t.setStyle(TableStyle(style))
    return t


def _pdf_payees(rep):
    widths = [22, CONTENT_W - 22 - 130 - 40 - 80 - 85, 130, 40, 80, 85]
    rows = [[P("#", "head"), P("ผู้รับ", "head"), P("หมวดหลัก", "head"), P("รายการ", "head_r"),
             P("ยอดรวม", "head_r"), P("สัดส่วน", "head")]]
    for i, p in enumerate(rep["top_payees"], start=1):
        share = _share_cell(p["share"])
        rows.append([P(i, "cell_muted"), PW(p["name"], widths[1]), PW(p["main_category"], widths[2], "cell_muted"),
                     P(p["count"], "cell_r"), P(fmt_baht(p["total"]), "cell_r", bold=True), share])
    t = Table(rows, colWidths=widths, repeatRows=1)
    t.setStyle(TableStyle(_base_table_style() + [("VALIGN", (5, 0), (5, -1), "MIDDLE")]))
    return t


def _pdf_largest(rep):
    widths = [62, 150, CONTENT_W - 62 - 150 - 110 - 75, 110, 75]
    rows = [[P("วันที่", "head"), P("ผู้รับ", "head"), P("รายละเอียด", "head"), P("หมวด", "head"), P("ยอดเงิน", "head_r")]]
    for t in rep["largest"]:
        rows.append([P(fmt_date(t["date"]), "cell_muted"), PW(t["receiver"] or UNKNOWN_PAYEE, widths[1]),
                     PW(t["memo"], widths[2], "cell_muted", max_lines=2), PW(t["category"], widths[3], "cell_muted"),
                     P(fmt_baht(t["amount"]), "cell_r", bold=True)])
    tbl = Table(rows, colWidths=widths, repeatRows=1)
    tbl.setStyle(TableStyle(_base_table_style()))
    return tbl


TXN_W = [50, 30, 140, None, 72, 66]


def _txn_widths():
    w = list(TXN_W)
    w[3] = CONTENT_W - sum(x for x in w if x)
    return w


def _pdf_txn_rows(items, widths, show_category=False):
    rows = []
    for t in items:
        detail = t["memo"]
        if show_category:
            detail = f"[{t['category']}] {detail}".strip()
        rows.append([
            P(fmt_date(t["date"], year=False), "cell_muted"), P(t["time"], "cell_muted"),
            PW(t["receiver"] or UNKNOWN_PAYEE, widths[2]), PW(detail, widths[3], "cell_muted"),
            P(t["bank"], "cell_muted"), P(fmt_baht(t["amount"]), "cell_r"),
        ])
    return rows


def _pdf_category_detail(rep):
    widths = _txn_widths()
    head = [P("วันที่", "head"), P("เวลา", "head"), P("ผู้รับ", "head"), P("รายละเอียด", "head"),
            P("ธนาคาร", "head"), P("ยอดเงิน", "head_r")]
    out = []
    for g, c, items in txns_by_category(rep):
        group_note = "" if c["name"] == g["name"] else f'  <font size="8" color="#{C_MUTED}">· {escape(g["name"])}</font>'
        title = Table([[Paragraph(f"<b>{escape(c['name'])}</b>{group_note}", ST["h3"]),
                        Paragraph(f'<font color="#{C_MUTED}" size="8">{len(items)} รายการ</font>  '
                                  f'<b>{fmt_baht(c["cur"])}</b>', _ps("ct", parent=ST["h3"], alignment=TA_RIGHT))]],
                      colWidths=[CONTENT_W * 0.65, CONTENT_W * 0.35])
        title.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), _hex(C_ACCENT_SOFT)),
            ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        body = Table([head] + _pdf_txn_rows(items, widths), colWidths=widths, repeatRows=1)
        body.setStyle(TableStyle(_base_table_style()))
        out += [CondPageBreak(3 * cm), KeepTogether([title, body]) if len(items) <= 12 else title]
        if len(items) > 12:
            out.append(body)
        out.append(Spacer(1, 10))
    return out


def _pdf_review(rep):
    widths = [56, 150, CONTENT_W - 56 - 150 - 100 - 70, 100, 70]
    rows = [[P("วันที่", "head"), P("ผู้รับ", "head"), P("รายละเอียด", "head"), P("ปัญหา", "head"), P("ยอดเงิน", "head_r")]]
    for t in rep["review"]:
        rows.append([P(fmt_date(t["date"], year=False), "cell_muted"), PW(t["receiver"] or UNKNOWN_PAYEE, widths[1]),
                     PW(t["memo"], widths[2], "cell_muted"), P(t["issue"], "cell", color=C_WARN),
                     P(fmt_baht(t["amount"]), "cell_r")])
    tbl = Table(rows, colWidths=widths, repeatRows=1)
    tbl.setStyle(TableStyle(_base_table_style()))
    return tbl


def build_report_pdf(rep, output):
    """Render a fetch_period_report() dict to an A4 PDF. `output` is a path or a file-like object."""
    k = rep["kpi"]
    story = [_pdf_header(rep), Spacer(1, 12), _pdf_kpis(rep), Spacer(1, 10)]
    if rep["partial"]:
        story += [_callout(f"ช่วงนี้ยังไม่จบ — ข้อมูลถึง {fmt_date(rep['today'])} "
                           f"ยอดที่เทียบกับ {escape(rep['prev_label'])} จึงยังไม่เต็มช่วง"), Spacer(1, 6)]
    if k["review_count"]:
        parts = []
        if k["uncategorized_count"]:
            parts.append(f"ยังไม่ระบุหมวด <b>{k['uncategorized_count']}</b> รายการ")
        if k["unknown_count"]:
            parts.append(f"ไม่แน่ใจว่าเป็นรายจ่าย <b>{k['unknown_count']}</b> รายการ (ไม่รวมในยอด)")
        story += [_callout("มีรายการต้องตรวจ: " + " · ".join(parts) + " — ดูรายละเอียดท้ายรายงาน"), Spacer(1, 6)]

    story += [Paragraph("สรุปสำคัญ", ST["h2"]), *_pdf_highlights(rep)]

    pie = _pdf_pie(rep)
    if pie:
        story += [KeepTogether([Paragraph("สัดส่วนรายจ่ายตามกลุ่ม", ST["h2"]), pie])]
    trend = _pdf_trend(rep)
    if trend:
        title = "รายจ่ายรายวัน" if rep["trend_unit"] == "day" else "รายจ่ายรายเดือน"
        story += [KeepTogether([Paragraph(title, ST["h2"]), trend])]

    story += [CondPageBreak(5 * cm), Paragraph("แยกตามกลุ่มและหมวด", ST["h2"]),
              Paragraph(f"เทียบกับ {escape(rep['prev_label'])} · สีแดง = จ่ายเพิ่มขึ้น, สีเขียว = จ่ายลดลง", ST["muted"]),
              Spacer(1, 4), _pdf_breakdown(rep)]
    if rep["top_payees"]:
        story += [CondPageBreak(5 * cm), Paragraph("ผู้รับเงินสูงสุด 10 อันดับ", ST["h2"]), _pdf_payees(rep)]
    if rep["largest"]:
        story += [CondPageBreak(5 * cm), Paragraph("รายการใหญ่ที่สุด 10 อันดับ", ST["h2"]), _pdf_largest(rep)]

    if rep["txns"]:
        story += [PageBreak(), Paragraph("รายละเอียดรายการตามหมวด", ST["h2"]),
                  Paragraph(f"ทั้งหมด {k['count']} รายการ รวม {fmt_baht(k['total'])} บาท", ST["muted"]),
                  Spacer(1, 6), *_pdf_category_detail(rep)]
    if rep["review"]:
        story += [CondPageBreak(5 * cm), Paragraph("รายการที่ต้องตรวจ", ST["h2"]),
                  Paragraph("รายการที่ 'ไม่แน่ใจว่าเป็นรายจ่าย' ไม่ได้รวมในยอดข้างต้น", ST["muted"]),
                  Spacer(1, 4), _pdf_review(rep)]
    if not rep["txns"] and not rep["review"]:
        story += [Spacer(1, 20), P("ไม่มีรายการในช่วงนี้", "muted")]

    footer = f"{REPORT_TITLE} · {rep['label']} · จัดทำ {fmt_date(rep['generated_at'].date())}"
    doc = SimpleDocTemplate(
        output, pagesize=A4, title=f"{REPORT_TITLE} {rep['label']}", author="Ledger",
        topMargin=MARGIN, bottomMargin=MARGIN + 0.3 * cm, leftMargin=MARGIN, rightMargin=MARGIN,
    )
    doc.build(story, canvasmaker=lambda *a, **kw: _NumberedCanvas(*a, footer_left=footer, **kw))


# ---------------------------------------------------------------------------
# Ledger listing PDF (dashboard filter export)
# ---------------------------------------------------------------------------

def build_pdf(df, output):
    """Render a filtered transaction DataFrame to a landscape PDF listing."""
    page_w = landscape(A4)[0] - 2 * 1.2 * cm
    header = list(df.columns)
    rows = df.fillna("").astype(str).values.tolist()
    weights = {"รายละเอียด": 3, "ผู้รับ": 2.2, "ผู้โอน": 2.2, "หมวด": 1.8}
    raw = [weights.get(h, 1) for h in header]
    widths = [page_w * w / sum(raw) for w in raw]
    money = {"ยอดเงิน", "ค่าธรรมเนียม"}
    table_data = [[P(h, "head_r" if h in money else "head") for h in header]]
    for r in rows:
        cells = []
        for h, v, w in zip(header, r, widths):
            if h in money and v:
                cells.append(P(fmt_baht(float(v)), "cell_r"))
            elif h == "เวลา":
                cells.append(P(v[:5], "cell_muted"))
            else:
                cells.append(PW(v, w))
        table_data.append(cells)

    t = Table(table_data, colWidths=widths, repeatRows=1)
    t.setStyle(TableStyle(_base_table_style() + [
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, _hex(C_SURFACE_ALT)]),
    ]))

    total_expense = df.loc[df["ประเภท"] == "expense", "ยอดเงิน"].sum()

    story = [
        Paragraph("รายการสลิปธนาคาร", ST["h2"]),
        Paragraph(f"จำนวน {len(df)} รายการ · รวมรายจ่าย {total_expense:,.2f} บาท", ST["body"]),
        Spacer(1, 0.3 * cm),
        t,
    ]
    footer = f"รายการสลิปธนาคาร · จัดทำ {fmt_date(datetime.now(BANGKOK_TZ).date())}"
    SimpleDocTemplate(
        output, pagesize=landscape(A4),
        topMargin=1.5 * cm, bottomMargin=1.8 * cm, leftMargin=1.2 * cm, rightMargin=1.2 * cm,
    ).build(story, canvasmaker=lambda *a, **kw: _LandscapeCanvas(*a, footer_left=footer, **kw))


class _LandscapeCanvas(_NumberedCanvas):
    def _draw_footer(self, total):
        w = landscape(A4)[0]
        m = 1.2 * cm
        y = 0.9 * cm
        self.setStrokeColor(_hex(C_BORDER))
        self.setLineWidth(0.5)
        self.line(m, y + 11, w - m, y + 11)
        self.setFont(FONT, 7.5)
        self.setFillColor(_hex(C_MUTED))
        self.drawString(m, y, self._footer_left, shaping=SHAPING)
        self.drawRightString(w - m, y, f"หน้า {self._pageNumber} / {total}", shaping=SHAPING)


def main():
    import psycopg2
    from dotenv import load_dotenv

    load_dotenv()
    today = datetime.now(BANGKOK_TZ).date()
    if len(sys.argv) > 1:
        y, m = (int(p) for p in sys.argv[1].split("-"))
    else:
        y, m = today.year, today.month
    d_from = date(y, m, 1)
    conn = psycopg2.connect(os.environ["DATABASE_URL"])
    try:
        with conn.cursor() as cur:
            rep = fetch_period_report(cur, d_from, _month_end(d_from))
    finally:
        conn.close()
    for ext, build in (("xlsx", build_report_xlsx), ("pdf", build_report_pdf)):
        path = report_filename(rep, ext)
        build(rep, path)
        print(f"-> {path}")


if __name__ == "__main__":
    main()
