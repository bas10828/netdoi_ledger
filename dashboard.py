"""Web dashboard for slip_ledger — login + transaction table, presentable for management.

Run:
  pip install -r requirements.txt
  fill in DATABASE_URL / DASHBOARD_SESSION_SECRET in .env
  uvicorn dashboard:app --host 0.0.0.0 --port 8081
"""

import io
import json
import os
import uuid
from collections import defaultdict
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path

import bcrypt
import pandas as pd
import psycopg2
import psycopg2.extras
from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, RedirectResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from linebot.v3.messaging import ApiClient, Configuration, MessagingApi
from starlette.middleware.sessions import SessionMiddleware

from categorize import get_categories, guess_category

load_dotenv()

LINE_FILES_DIR = Path("line_files")
LINE_FILES_DIR.mkdir(exist_ok=True)

DATABASE_URL = os.environ["DATABASE_URL"]
SESSION_SECRET = os.environ["DASHBOARD_SESSION_SECRET"]
LINE_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")

THAI_MONTHS = {
    1: "ม.ค.", 2: "ก.พ.", 3: "มี.ค.", 4: "เม.ย.", 5: "พ.ค.", 6: "มิ.ย.",
    7: "ก.ค.", 8: "ส.ค.", 9: "ก.ย.", 10: "ต.ค.", 11: "พ.ย.", 12: "ธ.ค.",
}

app = FastAPI()
app.add_middleware(SessionMiddleware, secret_key=SESSION_SECRET, same_site="lax")
templates = Jinja2Templates(directory="templates")
templates.env.filters["baht"] = lambda v: f"{v:,.2f}"


def db():
    return psycopg2.connect(DATABASE_URL)


def require_login(request: Request):
    if not request.session.get("user"):
        return RedirectResponse("/login")
    return None


def require_admin(request: Request):
    redirect = require_login(request)
    if redirect:
        return redirect
    if request.session.get("role") != "admin":
        raise HTTPException(403, "ต้องเป็น admin ถึงแก้ไขได้")
    return None


@app.get("/login")
def login_form(request: Request):
    if request.session.get("user"):
        return RedirectResponse("/")
    return templates.TemplateResponse(request, "login.html", {"error": None})


@app.post("/login")
def login_submit(request: Request, username: str = Form(...), password: str = Form(...)):
    conn = db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT password_hash, role FROM dashboard_users WHERE username = %s", (username,))
            row = cur.fetchone()
    finally:
        conn.close()

    if not row or not bcrypt.checkpw(password.encode(), row[0].encode()):
        return templates.TemplateResponse(
            request, "login.html", {"error": "username/password ไม่ถูกต้อง"}, status_code=401
        )

    request.session["user"] = username
    request.session["role"] = row[1]
    return RedirectResponse("/", status_code=302)


@app.get("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login")


@app.get("/files/{raw_file_id}")
def serve_file(raw_file_id: int, request: Request):
    redirect = require_login(request)
    if redirect:
        return redirect

    conn = db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT storage_path FROM raw_files WHERE id = %s", (raw_file_id,))
            row = cur.fetchone()
    finally:
        conn.close()

    if not row:
        raise HTTPException(404)
    return FileResponse(row[0])


def get_setting(cur, key, default=None):
    cur.execute("SELECT value FROM app_settings WHERE key = %s", (key,))
    row = cur.fetchone()
    return row["value"] if row else default


def log_audit(cur, txn_id, action, username, before=None, after=None):
    cur.execute(
        """INSERT INTO audit_log (txn_id, action, changed_by, before_data, after_data)
           VALUES (%s, %s, %s, %s, %s)""",
        (
            txn_id,
            action,
            username,
            json.dumps(before, default=str, ensure_ascii=False) if before is not None else None,
            json.dumps(after, default=str, ensure_ascii=False) if after is not None else None,
        ),
    )


def save_uploaded_photo(cur, photo):
    if not photo or not photo.filename:
        return None
    ext = Path(photo.filename).suffix or ".jpg"
    local_path = LINE_FILES_DIR / f"{uuid.uuid4().hex}{ext}"
    local_path.write_bytes(photo.file.read())
    cur.execute(
        """INSERT INTO raw_files (file_type, storage_path, is_slip, processed)
           VALUES ('image', %s, false, true) RETURNING id""",
        (str(local_path),),
    )
    return cur.fetchone()[0]


def shell_context(cur, request):
    """Template variables the app shell (base.html sidebar) needs on every page.
    Works regardless of the calling cursor's factory."""
    with cur.connection.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as c:
        c.execute("SELECT COALESCE(SUM(cost_usd), 0) AS total FROM ai_usage_log")
        ai_spent = float(c.fetchone()["total"])
        ai_starting_balance = float(get_setting(c, "ai_starting_balance_usd", "0"))
    return {
        "user": request.session.get("user"),
        "role": request.session.get("role"),
        "ai_balance": ai_starting_balance - ai_spent,
        "ai_starting_balance": ai_starting_balance,
    }


def render_page(request, template, nav, context, status_code=200):
    """Render a page that extends base.html — adds the shell's sidebar context."""
    conn = db()
    try:
        with conn.cursor() as cur:
            shell = shell_context(cur, request)
    finally:
        conn.close()
    return templates.TemplateResponse(request, template, {**shell, "nav": nav, **context}, status_code=status_code)


@app.get("/")
def dashboard(request: Request):
    redirect = require_login(request)
    if redirect:
        return redirect

    conn = db()
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                """SELECT id, raw_file_id, txn_date, txn_time, direction, category, amount, fee, bank,
                          sender_name, sender_account, receiver_name, receiver_account, memo,
                          ai_model, verified_bank, status, qr_trans_ref, printed_ref
                   FROM slip_transactions
                   ORDER BY txn_date DESC, txn_time DESC, id DESC"""
            )
            rows = cur.fetchall()
            categories = get_categories(cur)
            shell = shell_context(cur, request)
    finally:
        conn.close()

    groups = defaultdict(list)
    for r in rows:
        dup_key = r["qr_trans_ref"] or r["printed_ref"]
        if dup_key:
            groups[dup_key].append(r)
    dup_group = {}
    dup_extra_ids = set()
    for i, group in enumerate(g for g in groups.values() if len(g) > 1):
        for r in group:
            dup_group[r["id"]] = i + 1
        canonical = min(
            group, key=lambda r: (r["txn_date"] or date.min, r["txn_time"] or time.min, r["id"])
        )
        dup_extra_ids.update(r["id"] for r in group if r["id"] != canonical["id"])

    txns = [
        {
            "id": r["id"], "file": r["raw_file_id"],
            "date": r["txn_date"].isoformat() if r["txn_date"] else "",
            "time": r["txn_time"].strftime("%H:%M") if r["txn_time"] else "",
            "direction": r["direction"], "category": r["category"] or "",
            "amount": float(r["amount"]), "fee": float(r["fee"] or 0), "bank": r["bank"] or "",
            "sender": r["sender_name"] or "", "sender_acct": r["sender_account"] or "",
            "receiver": r["receiver_name"] or "", "receiver_acct": r["receiver_account"] or "",
            "memo": r["memo"] or "", "model": r["ai_model"] or "", "verified": bool(r["verified_bank"]),
            "status": r["status"], "ref": r["qr_trans_ref"] or r["printed_ref"] or "",
            "dup": dup_group.get(r["id"], 0), "dup_extra": r["id"] in dup_extra_ids,
        }
        for r in rows
    ]

    return templates.TemplateResponse(
        request,
        "dashboard.html",
        {
            **shell,
            "nav": "txns",
            # Embedded in a <script> tag and carries free-text memos — escape "<".
            "txns_json": json.dumps(txns, ensure_ascii=False).replace("<", "\\u003c"),
            "categories_json": json.dumps(
                [{"name": c["name"], "group": c["group_name"] or ""} for c in categories], ensure_ascii=False
            ).replace("<", "\\u003c"),
        },
    )


BANGKOK_TZ = timezone(timedelta(hours=7))
UNCATEGORIZED = "ไม่ระบุหมวด"
PERIOD_KEYS = ("cur", "prev", "ytd", "all")


def _month_bounds(y, m):
    start = date(y, m, 1)
    next_start = date(y + 1, 1, 1) if m == 12 else date(y, m + 1, 1)
    return start, next_start - timedelta(days=1)


def _shift_month(y, m, delta):
    idx = y * 12 + (m - 1) + delta
    return idx // 12, idx % 12 + 1


def _month_label(y, m):
    return f"{THAI_MONTHS[m]} {y}"


def _pct_change(cur, prev):
    if not prev:
        return None
    return (cur - prev) / prev * 100


def fetch_overview(cur, y, m):
    """Everything the /reports overview needs, anchored on month (y, m):
    month KPIs, top payees, and one drill-down "view" per category group (plus an
    all-expenses view). Each view carries its members'
    this-month / last-month / year-to-date / all-time totals, a 12-month trend,
    and this month's transactions, so the page can switch views client-side."""
    cur_from, cur_to = _month_bounds(y, m)
    prev_from, prev_to = _month_bounds(*_shift_month(y, m, -1))
    ytd_from = date(y, 1, 1)
    trend_from = _month_bounds(*_shift_month(y, m, -11))[0]

    categories = get_categories(cur)
    cat_order = {c["name"]: i for i, c in enumerate(categories)}
    cat_group = {c["name"]: c["group_name"] or c["name"] for c in categories}

    cur.execute(
        """SELECT COALESCE(category, %(unc)s),
                  COALESCE(SUM(amount) FILTER (WHERE txn_date BETWEEN %(cf)s AND %(ct)s), 0),
                  COALESCE(SUM(amount) FILTER (WHERE txn_date BETWEEN %(pf)s AND %(pt)s), 0),
                  COALESCE(SUM(amount) FILTER (WHERE txn_date BETWEEN %(yf)s AND %(ct)s), 0),
                  COALESCE(SUM(amount), 0)
           FROM slip_transactions
           WHERE direction = 'expense'
           GROUP BY 1""",
        {"unc": UNCATEGORIZED, "cf": cur_from, "ct": cur_to, "pf": prev_from, "pt": prev_to, "yf": ytd_from},
    )
    cat_amounts = {r[0]: dict(zip(PERIOD_KEYS, map(float, r[1:]))) for r in cur.fetchall()}

    months = [_shift_month(y, m, -i) for i in range(11, -1, -1)]
    month_keys = [f"{my:04d}-{mm:02d}" for my, mm in months]
    cur.execute(
        """SELECT to_char(txn_date, 'YYYY-MM'), COALESCE(category, %s), SUM(amount)
           FROM slip_transactions
           WHERE direction = 'expense' AND txn_date BETWEEN %s AND %s
           GROUP BY 1, 2""",
        (UNCATEGORIZED, trend_from, cur_to),
    )
    trend_by_cat = defaultdict(lambda: [0.0] * len(month_keys))
    for mk, cat, total in cur.fetchall():
        trend_by_cat[cat][month_keys.index(mk)] = float(total)

    # Monthly average only counts months since the ledger started, so a young ledger
    # isn't diluted by empty months before the first slip.
    cur.execute("SELECT to_char(MIN(txn_date), 'YYYY-MM') FROM slip_transactions")
    first_key = cur.fetchone()[0] or month_keys[-1]
    avg_months = sum(1 for k in month_keys if k >= first_key) or 1

    cur.execute(
        """SELECT id, txn_date, COALESCE(category, %s), receiver_name, memo, amount
           FROM slip_transactions
           WHERE direction = 'expense' AND txn_date BETWEEN %s AND %s
           ORDER BY txn_date DESC, txn_time DESC NULLS LAST, id DESC""",
        (UNCATEGORIZED, cur_from, cur_to),
    )
    month_txns = [
        {"id": r[0], "date": r[1].isoformat(), "category": r[2], "receiver": r[3] or "", "memo": r[4] or "",
         "amount": float(r[5])}
        for r in cur.fetchall()
    ]

    def build_view(key, label, cats, member_of, link_members=False):
        cats = sorted(cats, key=lambda n: cat_order.get(n, len(cat_order)))
        members = {}
        for c in cats:
            name = member_of(c)
            mbr = members.setdefault(
                name, {"name": name, "trend": [0.0] * len(month_keys), **{k: 0.0 for k in PERIOD_KEYS}}
            )
            for k in PERIOD_KEYS:
                mbr[k] += cat_amounts[c][k]
            mbr["trend"] = [a + b for a, b in zip(mbr["trend"], trend_by_cat.get(c, mbr["trend"]))]
        total = {k: sum(mb[k] for mb in members.values()) for k in PERIOD_KEYS}
        total["change"] = _pct_change(total["cur"], total["prev"])
        total["avg"] = sum(sum(mb["trend"]) for mb in members.values()) / avg_months
        for mb in members.values():
            mb["change"] = _pct_change(mb["cur"], mb["prev"])
            mb["link"] = f"g:{mb['name']}" if link_members else None
        cat_set = set(cats)
        return {
            "key": key, "label": label, "total": total,
            "members": sorted(members.values(), key=lambda mb: (-mb["cur"], -mb["all"])),
            "txns": [t for t in month_txns if t["category"] in cat_set],
        }

    all_cats = list(cat_amounts)
    views = [build_view("all", "รายจ่ายทั้งหมด", all_cats, lambda c: cat_group.get(c, c), link_members=True)]
    group_views = []
    for gname in {cat_group.get(c, c) for c in all_cats}:
        group_views.append(build_view(
            f"g:{gname}", gname, [c for c in all_cats if cat_group.get(c, c) == gname], lambda c: c,
        ))
    views += sorted(group_views, key=lambda v: -v["total"]["all"])

    cur.execute(
        """SELECT COALESCE(SUM(amount) FILTER (WHERE direction = 'expense' AND txn_date BETWEEN %(cf)s AND %(ct)s), 0),
                  COALESCE(SUM(amount) FILTER (WHERE direction = 'income' AND txn_date BETWEEN %(cf)s AND %(ct)s), 0),
                  COALESCE(SUM(amount) FILTER (WHERE direction = 'expense' AND txn_date BETWEEN %(pf)s AND %(pt)s), 0),
                  COALESCE(SUM(amount) FILTER (WHERE direction = 'income' AND txn_date BETWEEN %(pf)s AND %(pt)s), 0),
                  COUNT(*) FILTER (WHERE direction = 'unknown' AND txn_date BETWEEN %(cf)s AND %(ct)s),
                  COALESCE(SUM(amount) FILTER (WHERE direction = 'unknown' AND txn_date BETWEEN %(cf)s AND %(ct)s), 0),
                  COUNT(*) FILTER (WHERE direction = 'expense' AND txn_date BETWEEN %(cf)s AND %(ct)s)
           FROM slip_transactions""",
        {"cf": cur_from, "ct": cur_to, "pf": prev_from, "pt": prev_to},
    )
    exp_cur, inc_cur, exp_prev, inc_prev, unknown_count, unknown_total, expense_count = cur.fetchone()
    exp_cur, inc_cur, exp_prev, inc_prev = map(float, (exp_cur, inc_cur, exp_prev, inc_prev))
    kpis = {
        "expense": exp_cur, "income": inc_cur, "net": inc_cur - exp_cur,
        "expense_prev": exp_prev, "income_prev": inc_prev, "net_prev": inc_prev - exp_prev,
        "expense_change": _pct_change(exp_cur, exp_prev), "income_change": _pct_change(inc_cur, inc_prev),
        "expense_count": expense_count,
        "unknown_count": unknown_count, "unknown_total": float(unknown_total),
    }

    cur.execute(
        """SELECT COALESCE(receiver_name, 'ไม่ระบุ'), SUM(amount), COUNT(*)
           FROM slip_transactions
           WHERE direction = 'expense' AND txn_date BETWEEN %s AND %s
           GROUP BY 1 ORDER BY 2 DESC LIMIT 5""",
        (cur_from, cur_to),
    )
    top_payees = [{"name": r[0], "total": float(r[1]), "count": r[2]} for r in cur.fetchall()]

    return {
        "cur_from": cur_from, "cur_to": cur_to,
        "cur_label": _month_label(y, m), "prev_label": _month_label(*_shift_month(y, m, -1)),
        "ytd_label": f"ปี {y}",
        "month_labels": [f"{THAI_MONTHS[mm]} {str(my)[2:]}" for my, mm in months],
        "avg_months": avg_months, "views": views,
        "kpis": kpis, "top_payees": top_payees,
    }


@app.get("/reports")
def reports(request: Request, month: str = "", focus: str = "all"):
    redirect = require_login(request)
    if redirect:
        return redirect

    today = datetime.now(BANGKOK_TZ).date()
    try:
        y, m = (int(p) for p in month.split("-"))
        date(y, m, 1)
    except ValueError:
        y, m = today.year, today.month

    conn = db()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """SELECT DISTINCT EXTRACT(year FROM txn_date)::int, EXTRACT(month FROM txn_date)::int
                   FROM slip_transactions WHERE txn_date IS NOT NULL"""
            )
            month_opts = set(cur.fetchall()) | {(today.year, today.month), (y, m)}
            data = fetch_overview(cur, y, m)
    finally:
        conn.close()

    return render_page(
        request,
        "reports.html",
        "reports",
        {
            "selected_month": f"{y:04d}-{m:02d}",
            "month_options": [
                {"value": f"{oy:04d}-{om:02d}", "label": _month_label(oy, om)}
                for oy, om in sorted(month_opts, reverse=True)
            ],
            "focus": focus,
            # Embedded in a <script> tag and carries free-text memos — escape "<" so a
            # memo containing "</script>" can't break out of it.
            "report_json": json.dumps(
                {k: data[k] for k in ("views", "month_labels", "avg_months", "cur_label", "prev_label", "ytd_label")},
                ensure_ascii=False,
            ).replace("<", "\\u003c"),
            **data,
        },
    )


def _parse_report_range(date_from, date_to):
    try:
        d_from, d_to = date.fromisoformat(date_from), date.fromisoformat(date_to)
    except ValueError:
        raise HTTPException(400, "date_from / date_to ต้องเป็นวันที่รูปแบบ YYYY-MM-DD")
    if d_from > d_to:
        raise HTTPException(400, "date_from ต้องไม่เกิน date_to")
    return d_from, d_to


@app.get("/reports/export/{kind}")
def reports_export(request: Request, kind: str, date_from: str = "", date_to: str = ""):
    redirect = require_login(request)
    if redirect:
        return redirect
    if kind not in ("excel", "pdf"):
        raise HTTPException(404)
    d_from, d_to = _parse_report_range(date_from, date_to)

    try:
        import export_report
    except RuntimeError as e:
        raise HTTPException(500, str(e))

    conn = db()
    try:
        with conn.cursor() as cur:
            rep = export_report.fetch_period_report(cur, d_from, d_to)
    finally:
        conn.close()

    buf = io.BytesIO()
    if kind == "excel":
        export_report.build_report_xlsx(rep, buf)
        ext, media_type = "xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    else:
        export_report.build_report_pdf(rep, buf)
        ext, media_type = "pdf", "application/pdf"
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type=media_type,
        headers={"Content-Disposition": f'attachment; filename="{export_report.report_filename(rep, ext)}"'},
    )


def build_export_filter(direction, date_from, date_to, bank, status, name, category,
                         amount_min=None, amount_max=None):
    where = []
    params = []
    if category == "__none__":
        where.append("category IS NULL")
    elif category:
        where.append("category = %s")
        params.append(category)
    if direction:
        where.append("direction = ANY(%s)")
        params.append(direction.split(","))
    if date_from:
        where.append("txn_date >= %s")
        params.append(date_from)
    if date_to:
        where.append("txn_date <= %s")
        params.append(date_to)
    if amount_min is not None:
        where.append("amount >= %s")
        params.append(amount_min)
    if amount_max is not None:
        where.append("amount <= %s")
        params.append(amount_max)
    if bank:
        where.append("bank = %s")
        params.append(bank)
    if status == "verified":
        where.append("verified_bank = true")
    elif status == "unverified":
        where.append("verified_bank = false")
    elif status == "attention":
        where.append("(direction = 'unknown' OR (direction = 'expense' AND category IS NULL))")
    elif status == "dup":
        where.append("""COALESCE(qr_trans_ref, NULLIF(printed_ref, '')) IN (
            SELECT COALESCE(qr_trans_ref, NULLIF(printed_ref, '')) FROM slip_transactions
            WHERE COALESCE(qr_trans_ref, NULLIF(printed_ref, '')) IS NOT NULL
            GROUP BY COALESCE(qr_trans_ref, NULLIF(printed_ref, '')) HAVING COUNT(*) > 1
        )""")
    if name:
        where.append("(sender_name ILIKE %s OR receiver_name ILIKE %s OR memo ILIKE %s)")
        params.extend([f"%{name}%"] * 3)
    return where, params


def fetch_export_df(direction, date_from, date_to, bank, status, name, category,
                     amount_min=None, amount_max=None, columns="full"):
    where, params = build_export_filter(direction, date_from, date_to, bank, status, name, category,
                                         amount_min, amount_max)
    if columns == "print":
        select = """
            txn_date AS "วันที่", txn_time AS "เวลา", bank AS "ธนาคาร",
            direction AS "ประเภท", category AS "หมวด", amount AS "ยอดเงิน",
            receiver_name AS "ผู้รับ", memo AS "รายละเอียด"
        """
    else:
        select = """
            txn_date AS "วันที่", txn_time AS "เวลา", bank AS "ธนาคาร",
            direction AS "ประเภท", category AS "หมวด", amount AS "ยอดเงิน", fee AS "ค่าธรรมเนียม",
            sender_name AS "ผู้โอน", receiver_name AS "ผู้รับ", memo AS "รายละเอียด",
            qr_trans_ref AS "เลขอ้างอิง (QR)", verified_bank AS "ยืนยันธนาคาร"
        """
    query = f"SELECT {select} FROM slip_transactions"
    if where:
        query += " WHERE " + " AND ".join(where)
    query += " ORDER BY txn_date, txn_time"

    conn = db()
    try:
        return pd.read_sql(query, conn, params=params)
    finally:
        conn.close()


@app.get("/export/excel")
def export_excel(
    request: Request,
    direction: str = "",
    date_from: str = "",
    date_to: str = "",
    bank: str = "",
    status: str = "",
    name: str = "",
    category: str = "",
    amount_min: float | None = None,
    amount_max: float | None = None,
):
    redirect = require_login(request)
    if redirect:
        return redirect

    df = fetch_export_df(direction, date_from, date_to, bank, status, name, category,
                          amount_min, amount_max)

    summary = (
        df.groupby("ประเภท")["ยอดเงิน"]
        .sum()
        .reindex(["expense", "unknown"])
        .fillna(0)
        .rename({"expense": "รายจ่าย", "unknown": "ไม่ระบุประเภท"})
    )

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="รายการ", index=False)
        summary.to_frame("รวม (บาท)").to_excel(writer, sheet_name="สรุป")
    buf.seek(0)

    filename = f"ledger_{date.today().isoformat()}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.get("/export/pdf")
def export_pdf(
    request: Request,
    direction: str = "",
    date_from: str = "",
    date_to: str = "",
    bank: str = "",
    status: str = "",
    name: str = "",
    category: str = "",
    amount_min: float | None = None,
    amount_max: float | None = None,
):
    redirect = require_login(request)
    if redirect:
        return redirect

    try:
        from export_report import build_pdf
    except RuntimeError as e:
        raise HTTPException(500, str(e))

    df = fetch_export_df(direction, date_from, date_to, bank, status, name, category,
                          amount_min, amount_max, columns="print")

    buf = io.BytesIO()
    build_pdf(df, buf)
    buf.seek(0)

    filename = f"ledger_{date.today().isoformat()}.pdf"
    return StreamingResponse(
        buf,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/settings/ai-balance")
def update_ai_balance(request: Request, value: float = Form(...)):
    redirect = require_admin(request)
    if redirect:
        return redirect

    conn = db()
    try:
        with conn, conn.cursor() as cur:
            cur.execute(
                """INSERT INTO app_settings (key, value) VALUES ('ai_starting_balance_usd', %s)
                   ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value""",
                (str(value),),
            )
    finally:
        conn.close()

    return RedirectResponse("/", status_code=302)


@app.get("/settings/categories")
def categories_view(request: Request):
    redirect = require_admin(request)
    if redirect:
        return redirect

    conn = db()
    try:
        with conn.cursor() as cur:
            categories = get_categories(cur)
    finally:
        conn.close()

    return render_page(request, "settings_categories.html", "categories", {"categories": categories})


@app.post("/settings/categories/new")
def categories_new(request: Request, name: str = Form(...), keywords: str = Form(""), group_name: str = Form("")):
    redirect = require_admin(request)
    if redirect:
        return redirect

    kw_list = [k.strip() for k in keywords.split(",") if k.strip()]
    conn = db()
    try:
        with conn, conn.cursor() as cur:
            cur.execute("SELECT COALESCE(MAX(sort_order), 0) + 1 FROM categories")
            next_order = cur.fetchone()[0]
            cur.execute(
                "INSERT INTO categories (name, keywords, sort_order, group_name) VALUES (%s, %s, %s, %s)",
                (name.strip(), kw_list, next_order, group_name.strip() or None),
            )
    finally:
        conn.close()

    return RedirectResponse("/settings/categories", status_code=302)


@app.post("/settings/categories/{cat_id}/edit")
def categories_edit(
    cat_id: int, request: Request, name: str = Form(...), keywords: str = Form(""), group_name: str = Form("")
):
    redirect = require_admin(request)
    if redirect:
        return redirect

    kw_list = [k.strip() for k in keywords.split(",") if k.strip()]
    conn = db()
    try:
        with conn, conn.cursor() as cur:
            cur.execute(
                "UPDATE categories SET name = %s, keywords = %s, group_name = %s WHERE id = %s",
                (name.strip(), kw_list, group_name.strip() or None, cat_id),
            )
    finally:
        conn.close()

    return RedirectResponse("/settings/categories", status_code=302)


@app.post("/settings/categories/{cat_id}/delete")
def categories_delete(cat_id: int, request: Request):
    redirect = require_admin(request)
    if redirect:
        return redirect

    conn = db()
    try:
        with conn, conn.cursor() as cur:
            cur.execute("DELETE FROM categories WHERE id = %s", (cat_id,))
    finally:
        conn.close()

    return RedirectResponse("/settings/categories", status_code=302)


@app.post("/settings/categories/{cat_id}/move")
def categories_move(cat_id: int, request: Request, direction: str = Form(...)):
    redirect = require_admin(request)
    if redirect:
        return redirect
    if direction not in ("up", "down"):
        raise HTTPException(400)

    conn = db()
    try:
        with conn, conn.cursor() as cur:
            cur.execute("SELECT id, sort_order FROM categories ORDER BY sort_order, id")
            rows = cur.fetchall()
            ids = [r[0] for r in rows]
            idx = ids.index(cat_id)
            swap_idx = idx - 1 if direction == "up" else idx + 1
            if 0 <= swap_idx < len(rows):
                a_id, a_order = rows[idx]
                b_id, b_order = rows[swap_idx]
                cur.execute("UPDATE categories SET sort_order = %s WHERE id = %s", (b_order, a_id))
                cur.execute("UPDATE categories SET sort_order = %s WHERE id = %s", (a_order, b_id))
    finally:
        conn.close()

    return RedirectResponse("/settings/categories", status_code=302)


def resolve_display_names(messages):
    names = {}
    if not LINE_ACCESS_TOKEN:
        return names
    config = Configuration(access_token=LINE_ACCESS_TOKEN)
    with ApiClient(config) as api_client:
        api = MessagingApi(api_client)
        for m in messages:
            user_id = m.get("line_user_id")
            group_id = m.get("line_group_id")
            if not user_id or not group_id or user_id in names:
                continue
            try:
                profile = api.get_group_member_profile(group_id, user_id)
                names[user_id] = profile.display_name
            except Exception:
                pass
    return names


AUDIT_FIELD_LABELS = {
    "bank": "ธนาคาร", "txn_date": "วันที่", "txn_time": "เวลา", "amount": "ยอดเงิน", "fee": "ค่าธรรมเนียม",
    "sender_name": "ผู้โอน", "receiver_name": "ผู้รับ", "memo": "รายละเอียด",
    "direction": "ประเภท", "category": "หมวด", "status": "สถานะ",
}
AUDIT_ACTION_LABELS = {"create": "เพิ่มรายการ", "update": "แก้ไขรายการ", "delete": "ลบรายการ"}


def _normalize_for_diff(key, value):
    if value is None:
        return None
    if key in ("amount", "fee"):
        try:
            return f"{float(value):.2f}"
        except (TypeError, ValueError):
            return str(value)
    if key == "txn_time":
        return str(value)[:5]
    return str(value)


def summarize_audit_entry(action, before, after):
    if action == "create":
        return f"{(after or {}).get('memo') or '-'} ฿{(after or {}).get('amount', '-')}"
    if action == "delete":
        return f"{(before or {}).get('memo') or '-'} ฿{(before or {}).get('amount', '-')}"
    if not before or not after:
        return "-"
    changes = []
    for key, label in AUDIT_FIELD_LABELS.items():
        old_val, new_val = before.get(key), after.get(key)
        if _normalize_for_diff(key, old_val) != _normalize_for_diff(key, new_val):
            changes.append(f"{label}: {'ไม่ระบุ' if old_val is None else old_val} → {'ไม่ระบุ' if new_val is None else new_val}")
    return "; ".join(changes) if changes else "ไม่มีการเปลี่ยนแปลง"


@app.get("/audit")
def audit_view(request: Request):
    redirect = require_admin(request)
    if redirect:
        return redirect

    conn = db()
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                """SELECT id, txn_id, action, changed_by, changed_at, before_data, after_data
                   FROM audit_log
                   ORDER BY changed_at DESC
                   LIMIT 500"""
            )
            entries = cur.fetchall()
    finally:
        conn.close()

    for e in entries:
        e["action_label"] = AUDIT_ACTION_LABELS.get(e["action"], e["action"])
        e["summary"] = summarize_audit_entry(e["action"], e["before_data"], e["after_data"])

    return render_page(request, "audit.html", "audit", {"entries": entries})


@app.get("/chat")
def chat_view(request: Request):
    redirect = require_login(request)
    if redirect:
        return redirect

    conn = db()
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                """SELECT 'text' AS kind, id, line_user_id, line_group_id, text,
                          NULL::integer AS raw_file_id, received_at
                       FROM line_messages
                   UNION ALL
                   SELECT file_type AS kind, id, line_user_id, line_group_id, NULL AS text,
                          id AS raw_file_id, received_at
                       FROM raw_files
                   ORDER BY received_at DESC
                   LIMIT 500"""
            )
            messages = list(reversed(cur.fetchall()))
    finally:
        conn.close()

    names = resolve_display_names(messages)

    return render_page(request, "chat.html", "chat", {"messages": messages, "names": names})


@app.get("/transactions/new")
def new_form(request: Request):
    redirect = require_admin(request)
    if redirect:
        return redirect

    empty = {
        "id": None, "raw_file_id": None, "bank": "", "txn_date": "", "txn_time": "",
        "amount": "", "fee": 0, "sender_name": "", "sender_account": "",
        "receiver_name": "", "receiver_account": "", "memo": "",
        "direction": "expense", "category": "", "status": "reviewed",
    }
    conn = db()
    try:
        with conn.cursor() as cur:
            categories = [c["name"] for c in get_categories(cur)]
    finally:
        conn.close()

    return render_page(
        request, "edit_transaction.html", "txns",
        {"txn": empty, "form_action": "/transactions/new", "categories": categories},
    )


@app.post("/transactions/new")
def new_submit(
    request: Request,
    bank: str = Form(""),
    txn_date: str = Form(...),
    txn_time: str = Form(...),
    amount: float = Form(...),
    fee: float = Form(0),
    sender_name: str = Form(""),
    sender_account: str = Form(""),
    receiver_name: str = Form(""),
    receiver_account: str = Form(""),
    memo: str = Form(""),
    direction: str = Form("expense"),
    category: str = Form(""),
    status: str = Form("reviewed"),
    photo: UploadFile = File(None),
):
    redirect = require_admin(request)
    if redirect:
        return redirect

    conn = db()
    try:
        with conn, conn.cursor() as cur:
            final_category = category or guess_category(memo, get_categories(cur))
            raw_file_id = save_uploaded_photo(cur, photo)
            cur.execute(
                """INSERT INTO slip_transactions
                       (raw_file_id, bank, txn_date, txn_time, amount, fee, sender_name, sender_account,
                        receiver_name, receiver_account, memo, direction, category, status, ai_model)
                   VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
                (raw_file_id, bank, txn_date, txn_time, amount, fee, sender_name, sender_account,
                 receiver_name, receiver_account, memo, direction, final_category, status, "manual"),
            )
            new_id = cur.fetchone()[0]
            log_audit(
                cur, new_id, "create", request.session.get("user"),
                after={
                    "bank": bank, "txn_date": txn_date, "txn_time": txn_time, "amount": amount, "fee": fee,
                    "sender_name": sender_name, "receiver_name": receiver_name, "memo": memo,
                    "direction": direction, "category": final_category, "status": status,
                },
            )
    finally:
        conn.close()

    return RedirectResponse("/", status_code=302)


@app.get("/transactions/{txn_id}/edit")
def edit_form(txn_id: int, request: Request):
    redirect = require_admin(request)
    if redirect:
        return redirect

    conn = db()
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute("SELECT * FROM slip_transactions WHERE id = %s", (txn_id,))
            txn = cur.fetchone()
            categories = get_categories(cur)
    finally:
        conn.close()

    if not txn:
        raise HTTPException(404)
    if not txn["category"]:
        txn["category"] = guess_category(txn["memo"], categories)
    return render_page(
        request, "edit_transaction.html", "txns",
        {"txn": txn, "form_action": f"/transactions/{txn_id}/edit", "categories": [c["name"] for c in categories]},
    )


@app.post("/transactions/{txn_id}/edit")
def edit_submit(
    txn_id: int,
    request: Request,
    bank: str = Form(""),
    txn_date: str = Form(...),
    txn_time: str = Form(...),
    amount: float = Form(...),
    fee: float = Form(0),
    sender_name: str = Form(""),
    sender_account: str = Form(""),
    receiver_name: str = Form(""),
    receiver_account: str = Form(""),
    memo: str = Form(""),
    direction: str = Form("unknown"),
    category: str = Form(""),
    status: str = Form("reviewed"),
    photo: UploadFile = File(None),
):
    redirect = require_admin(request)
    if redirect:
        return redirect

    final_category = category or None
    conn = db()
    try:
        with conn, conn.cursor() as cur:
            cur.execute("SELECT * FROM slip_transactions WHERE id = %s", (txn_id,))
            columns = [c.name for c in cur.description]
            before_row = cur.fetchone()
            before = dict(zip(columns, before_row)) if before_row else None

            new_raw_file_id = save_uploaded_photo(cur, photo)
            raw_file_id = new_raw_file_id if new_raw_file_id is not None else (before or {}).get("raw_file_id")

            cur.execute(
                """UPDATE slip_transactions SET
                       raw_file_id=%s, bank=%s, txn_date=%s, txn_time=%s, amount=%s, fee=%s,
                       sender_name=%s, sender_account=%s, receiver_name=%s, receiver_account=%s,
                       memo=%s, direction=%s, category=%s, status=%s
                   WHERE id=%s""",
                (raw_file_id, bank, txn_date, txn_time, amount, fee, sender_name, sender_account,
                 receiver_name, receiver_account, memo, direction, final_category, status, txn_id),
            )
            log_audit(
                cur, txn_id, "update", request.session.get("user"),
                before=before,
                after={
                    "bank": bank, "txn_date": txn_date, "txn_time": txn_time, "amount": amount, "fee": fee,
                    "sender_name": sender_name, "receiver_name": receiver_name, "memo": memo,
                    "direction": direction, "category": final_category, "status": status,
                },
            )
    finally:
        conn.close()

    return RedirectResponse("/", status_code=302)


@app.post("/transactions/{txn_id}/direction")
def update_direction(txn_id: int, request: Request, direction: str = Form(...)):
    redirect = require_admin(request)
    if redirect:
        return redirect
    if direction not in ("expense", "income", "unknown"):
        raise HTTPException(400)

    conn = db()
    try:
        with conn, conn.cursor() as cur:
            cur.execute("SELECT direction FROM slip_transactions WHERE id = %s", (txn_id,))
            old = cur.fetchone()
            cur.execute("UPDATE slip_transactions SET direction = %s WHERE id = %s", (direction, txn_id))
            log_audit(
                cur, txn_id, "update", request.session.get("user"),
                before={"direction": old[0] if old else None}, after={"direction": direction},
            )
    finally:
        conn.close()

    return RedirectResponse("/", status_code=302)


@app.post("/transactions/{txn_id}/category")
def update_category(txn_id: int, request: Request, category: str = Form("")):
    redirect = require_admin(request)
    if redirect:
        return redirect

    final_category = category or None
    conn = db()
    try:
        with conn, conn.cursor() as cur:
            if category:
                valid_names = [c["name"] for c in get_categories(cur)]
                if category not in valid_names:
                    raise HTTPException(400)
            cur.execute("SELECT category FROM slip_transactions WHERE id = %s", (txn_id,))
            old = cur.fetchone()
            cur.execute("UPDATE slip_transactions SET category = %s WHERE id = %s", (final_category, txn_id))
            log_audit(
                cur, txn_id, "update", request.session.get("user"),
                before={"category": old[0] if old else None}, after={"category": final_category},
            )
    finally:
        conn.close()

    return RedirectResponse("/", status_code=302)


@app.post("/transactions/{txn_id}/delete")
def delete_submit(txn_id: int, request: Request):
    redirect = require_admin(request)
    if redirect:
        return redirect

    conn = db()
    try:
        with conn, conn.cursor() as cur:
            cur.execute("SELECT * FROM slip_transactions WHERE id = %s", (txn_id,))
            columns = [c.name for c in cur.description]
            before_row = cur.fetchone()
            before = dict(zip(columns, before_row)) if before_row else None

            cur.execute("DELETE FROM slip_transactions WHERE id = %s", (txn_id,))
            log_audit(cur, txn_id, "delete", request.session.get("user"), before=before)
    finally:
        conn.close()

    return RedirectResponse("/", status_code=302)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8081)
