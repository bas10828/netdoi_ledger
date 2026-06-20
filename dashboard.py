"""Web dashboard for slip_ledger — login + transaction table, presentable for management.

Run:
  pip install -r requirements.txt
  fill in DATABASE_URL / DASHBOARD_SESSION_SECRET in .env
  uvicorn dashboard:app --host 0.0.0.0 --port 8081
"""

import io
import json
import os
from collections import defaultdict
from datetime import date, time

import bcrypt
import pandas as pd
import psycopg2
import psycopg2.extras
from dotenv import load_dotenv
from fastapi import FastAPI, Form, HTTPException, Request
from fastapi.responses import FileResponse, RedirectResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from linebot.v3.messaging import ApiClient, Configuration, MessagingApi
from starlette.middleware.sessions import SessionMiddleware

from categorize import CATEGORIES, guess_category

load_dotenv()

DATABASE_URL = os.environ["DATABASE_URL"]
SESSION_SECRET = os.environ["DASHBOARD_SESSION_SECRET"]
LINE_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")

app = FastAPI()
app.add_middleware(SessionMiddleware, secret_key=SESSION_SECRET, same_site="lax")
templates = Jinja2Templates(directory="templates")
templates.env.globals["categories"] = CATEGORIES


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


@app.get("/")
def dashboard(request: Request):
    redirect = require_login(request)
    if redirect:
        return redirect

    conn = db()
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                """SELECT id, raw_file_id, txn_date, txn_time, direction, category, amount, bank,
                          sender_name, receiver_name, memo, ai_model, verified_bank, status,
                          qr_trans_ref, printed_ref
                   FROM slip_transactions
                   ORDER BY txn_date DESC, txn_time DESC, id DESC"""
            )
            rows = cur.fetchall()

            cur.execute("SELECT COALESCE(SUM(cost_usd), 0) AS total FROM ai_usage_log")
            ai_spent = float(cur.fetchone()["total"])
            ai_starting_balance = float(get_setting(cur, "ai_starting_balance_usd", "0"))
    finally:
        conn.close()

    groups = defaultdict(list)
    for r in rows:
        if r["qr_trans_ref"]:
            groups[r["qr_trans_ref"]].append(r)
    dup_color = {}
    dup_extra_ids = set()
    palette = ["dup-1", "dup-2", "dup-3", "dup-4", "dup-5"]
    i = 0
    for ref, group in groups.items():
        if len(group) > 1:
            for r in group:
                dup_color[r["id"]] = palette[i % len(palette)]
            i += 1
            canonical = min(
                group, key=lambda r: (r["txn_date"] or date.min, r["txn_time"] or time.min, r["id"])
            )
            dup_extra_ids.update(r["id"] for r in group if r["id"] != canonical["id"])

    counted_rows = [r for r in rows if r["id"] not in dup_extra_ids]
    expense = sum(float(r["amount"]) for r in counted_rows if r["direction"] == "expense")
    income = sum(float(r["amount"]) for r in counted_rows if r["direction"] == "income")

    return templates.TemplateResponse(
        request,
        "dashboard.html",
        {
            "user": request.session.get("user"),
            "role": request.session.get("role"),
            "rows": rows,
            "dup_color": dup_color,
            "expense": expense,
            "income": income,
            "net": income - expense,
            "count": len(counted_rows),
            "ai_balance": ai_starting_balance - ai_spent,
            "ai_starting_balance": ai_starting_balance,
        },
    )


@app.get("/reports")
def reports(request: Request, year: int = 0):
    redirect = require_login(request)
    if redirect:
        return redirect

    conn = db()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """SELECT DISTINCT EXTRACT(year FROM txn_date)::int AS y
                   FROM slip_transactions
                   WHERE txn_date IS NOT NULL
                   ORDER BY 1 DESC"""
            )
            years = [r[0] for r in cur.fetchall()]

            if not year:
                year = years[0] if years else date.today().year

            cur.execute(
                """SELECT to_char(date_trunc('month', txn_date), 'YYYY-MM') AS month,
                          COALESCE(SUM(amount) FILTER (WHERE direction = 'expense'), 0) AS expense,
                          COALESCE(SUM(amount) FILTER (WHERE direction = 'income'), 0) AS income
                   FROM slip_transactions
                   WHERE txn_date IS NOT NULL AND direction IN ('expense', 'income')
                         AND EXTRACT(year FROM txn_date) = %s
                   GROUP BY 1
                   ORDER BY 1""",
                (year,),
            )
            month_rows = cur.fetchall()

            cur.execute(
                """SELECT COALESCE(category, 'ไม่ระบุหมวด') AS category, SUM(amount) AS total
                   FROM slip_transactions
                   WHERE txn_date IS NOT NULL AND direction = 'expense'
                         AND EXTRACT(year FROM txn_date) = %s
                   GROUP BY 1
                   ORDER BY 2 DESC""",
                (year,),
            )
            category_rows = cur.fetchall()
    finally:
        conn.close()

    monthly_chart = {
        "labels": [r[0] for r in month_rows],
        "expense": [float(r[1]) for r in month_rows],
        "income": [float(r[2]) for r in month_rows],
    }
    category_chart = {
        "labels": [r[0] for r in category_rows],
        "amounts": [float(r[1]) for r in category_rows],
    }

    return templates.TemplateResponse(
        request,
        "reports.html",
        {
            "user": request.session.get("user"),
            "role": request.session.get("role"),
            "monthly_chart": json.dumps(monthly_chart),
            "category_chart": json.dumps(category_chart),
            "years": years,
            "selected_year": year,
        },
    )


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
):
    redirect = require_login(request)
    if redirect:
        return redirect

    where = []
    params = []
    if category:
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
    if bank:
        where.append("bank = %s")
        params.append(bank)
    if status == "verified":
        where.append("verified_bank = true")
    elif status == "unverified":
        where.append("verified_bank = false")
    elif status == "dup":
        where.append("""qr_trans_ref IN (
            SELECT qr_trans_ref FROM slip_transactions
            WHERE qr_trans_ref IS NOT NULL GROUP BY qr_trans_ref HAVING COUNT(*) > 1
        )""")
    if name:
        where.append("(sender_name ILIKE %s OR receiver_name ILIKE %s)")
        params.extend([f"%{name}%", f"%{name}%"])

    query = """
        SELECT
            txn_date AS "วันที่", txn_time AS "เวลา", bank AS "ธนาคาร",
            direction AS "ประเภท", category AS "หมวด", amount AS "ยอดเงิน", fee AS "ค่าธรรมเนียม",
            sender_name AS "ผู้โอน", receiver_name AS "ผู้รับ", memo AS "รายละเอียด",
            qr_trans_ref AS "เลขอ้างอิง (QR)", verified_bank AS "ยืนยันธนาคาร"
        FROM slip_transactions
    """
    if where:
        query += " WHERE " + " AND ".join(where)
    query += " ORDER BY txn_date, txn_time"

    conn = db()
    try:
        df = pd.read_sql(query, conn, params=params)
    finally:
        conn.close()

    summary = (
        df.groupby("ประเภท")["ยอดเงิน"]
        .sum()
        .reindex(["expense", "income", "unknown"])
        .fillna(0)
        .rename({"expense": "รายจ่าย", "income": "รายรับ", "unknown": "ไม่ระบุ"})
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
            changes.append(f"{label}: {old_val} → {new_val}")
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

    return templates.TemplateResponse(
        request,
        "audit.html",
        {"user": request.session.get("user"), "role": request.session.get("role"), "entries": entries},
    )


@app.get("/chat")
def chat_view(request: Request):
    redirect = require_admin(request)
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

    return templates.TemplateResponse(request, "chat.html", {"messages": messages, "names": names})


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
    return templates.TemplateResponse(
        request, "edit_transaction.html", {"txn": empty, "form_action": "/transactions/new"}
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
):
    redirect = require_admin(request)
    if redirect:
        return redirect

    final_category = category or guess_category(memo)
    conn = db()
    try:
        with conn, conn.cursor() as cur:
            cur.execute(
                """INSERT INTO slip_transactions
                       (bank, txn_date, txn_time, amount, fee, sender_name, sender_account,
                        receiver_name, receiver_account, memo, direction, category, status, ai_model)
                   VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
                (bank, txn_date, txn_time, amount, fee, sender_name, sender_account,
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
    finally:
        conn.close()

    if not txn:
        raise HTTPException(404)
    if not txn["category"]:
        txn["category"] = guess_category(txn["memo"])
    return templates.TemplateResponse(
        request, "edit_transaction.html", {"txn": txn, "form_action": f"/transactions/{txn_id}/edit"}
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

            cur.execute(
                """UPDATE slip_transactions SET
                       bank=%s, txn_date=%s, txn_time=%s, amount=%s, fee=%s,
                       sender_name=%s, sender_account=%s, receiver_name=%s, receiver_account=%s,
                       memo=%s, direction=%s, category=%s, status=%s
                   WHERE id=%s""",
                (bank, txn_date, txn_time, amount, fee, sender_name, sender_account,
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
    if category and category not in CATEGORIES:
        raise HTTPException(400)

    final_category = category or None
    conn = db()
    try:
        with conn, conn.cursor() as cur:
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
