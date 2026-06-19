"""Web dashboard for slip_ledger — login + transaction table, presentable for management.

Run:
  pip install -r requirements.txt
  fill in DATABASE_URL / DASHBOARD_SESSION_SECRET in .env
  uvicorn dashboard:app --host 0.0.0.0 --port 8081
"""

import os
from collections import defaultdict

import bcrypt
import psycopg2
import psycopg2.extras
from dotenv import load_dotenv
from fastapi import FastAPI, Form, HTTPException, Request
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from linebot.v3.messaging import ApiClient, Configuration, MessagingApi
from starlette.middleware.sessions import SessionMiddleware

load_dotenv()

DATABASE_URL = os.environ["DATABASE_URL"]
SESSION_SECRET = os.environ["DASHBOARD_SESSION_SECRET"]
LINE_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")

app = FastAPI()
app.add_middleware(SessionMiddleware, secret_key=SESSION_SECRET, same_site="lax")
templates = Jinja2Templates(directory="templates")


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


@app.get("/")
def dashboard(request: Request):
    redirect = require_login(request)
    if redirect:
        return redirect

    conn = db()
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                """SELECT id, raw_file_id, txn_date, txn_time, direction, amount, bank,
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

    expense = sum(float(r["amount"]) for r in rows if r["direction"] == "expense")
    income = sum(float(r["amount"]) for r in rows if r["direction"] == "income")

    groups = defaultdict(list)
    for r in rows:
        if r["qr_trans_ref"]:
            groups[r["qr_trans_ref"]].append(r["id"])
    dup_color = {}
    palette = ["dup-1", "dup-2", "dup-3", "dup-4", "dup-5"]
    i = 0
    for ref, ids in groups.items():
        if len(ids) > 1:
            for txn_id in ids:
                dup_color[txn_id] = palette[i % len(palette)]
            i += 1

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
            "count": len(rows),
            "ai_balance": ai_starting_balance - ai_spent,
            "ai_starting_balance": ai_starting_balance,
        },
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
        "direction": "expense", "status": "reviewed",
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
    status: str = Form("reviewed"),
):
    redirect = require_admin(request)
    if redirect:
        return redirect

    conn = db()
    try:
        with conn, conn.cursor() as cur:
            cur.execute(
                """INSERT INTO slip_transactions
                       (bank, txn_date, txn_time, amount, fee, sender_name, sender_account,
                        receiver_name, receiver_account, memo, direction, status, ai_model)
                   VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                (bank, txn_date, txn_time, amount, fee, sender_name, sender_account,
                 receiver_name, receiver_account, memo, direction, status, "manual"),
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
    status: str = Form("reviewed"),
):
    redirect = require_admin(request)
    if redirect:
        return redirect

    conn = db()
    try:
        with conn, conn.cursor() as cur:
            cur.execute(
                """UPDATE slip_transactions SET
                       bank=%s, txn_date=%s, txn_time=%s, amount=%s, fee=%s,
                       sender_name=%s, sender_account=%s, receiver_name=%s, receiver_account=%s,
                       memo=%s, direction=%s, status=%s
                   WHERE id=%s""",
                (bank, txn_date, txn_time, amount, fee, sender_name, sender_account,
                 receiver_name, receiver_account, memo, direction, status, txn_id),
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
            cur.execute("UPDATE slip_transactions SET direction = %s WHERE id = %s", (direction, txn_id))
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
            cur.execute("DELETE FROM slip_transactions WHERE id = %s", (txn_id,))
    finally:
        conn.close()

    return RedirectResponse("/", status_code=302)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8081)
