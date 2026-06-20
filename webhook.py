"""LINE webhook: backs up every file sent to the group, extracts slip data via Claude,
and stores both in Postgres.

Run:
  pip install -r requirements.txt
  fill in LINE_CHANNEL_SECRET / LINE_CHANNEL_ACCESS_TOKEN / DATABASE_URL in .env
  uvicorn webhook:app --host 0.0.0.0 --port 8000
  (separately) cloudflared tunnel --url http://localhost:8000
  set the printed https://*.trycloudflare.com/webhook URL as the LINE webhook URL
"""

import base64
import logging
import os
from pathlib import Path

import anthropic
import psycopg2
from dotenv import load_dotenv

from categorize import get_categories, guess_category
from fastapi import FastAPI, Header, HTTPException, Request
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.messaging import ApiClient, Configuration, MessagingApiBlob
from linebot.v3.webhook import WebhookParser
from linebot.v3.webhooks import (
    FileMessageContent,
    ImageMessageContent,
    MessageEvent,
    TextMessageContent,
)

from extract import MODEL, compute_cost, decode_qr, direction, extract_slip

load_dotenv()
logging.basicConfig(level=logging.INFO)
log = logging.getLogger("webhook")

CHANNEL_SECRET = os.environ["LINE_CHANNEL_SECRET"]
ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
DATABASE_URL = os.environ["DATABASE_URL"]
BACKUP_DIR = Path("line_files")
BACKUP_DIR.mkdir(exist_ok=True)

parser = WebhookParser(CHANNEL_SECRET)
claude = anthropic.Anthropic()
app = FastAPI()


def db():
    return psycopg2.connect(DATABASE_URL)


@app.post("/webhook")
async def webhook(request: Request, x_line_signature: str = Header(...)):
    body = (await request.body()).decode("utf-8")
    try:
        events = parser.parse(body, x_line_signature)
    except InvalidSignatureError:
        raise HTTPException(400, "invalid signature")

    for event in events:
        if not isinstance(event, MessageEvent):
            continue
        if isinstance(event.message, (ImageMessageContent, FileMessageContent)):
            handle_file(event)
        elif isinstance(event.message, TextMessageContent):
            handle_text(event)
    return "OK"


def handle_file(event: MessageEvent):
    msg = event.message
    is_image = isinstance(msg, ImageMessageContent)
    file_name = getattr(msg, "file_name", None) or f"{msg.id}.jpg"
    ext = Path(file_name).suffix or ".jpg"
    local_path = BACKUP_DIR / f"{msg.id}{ext}"

    if not ACCESS_TOKEN:
        log.warning("LINE_CHANNEL_ACCESS_TOKEN not set, skip download of %s", msg.id)
        return

    config = Configuration(access_token=ACCESS_TOKEN)
    with ApiClient(config) as api_client:
        content = MessagingApiBlob(api_client).get_message_content(msg.id)
    local_path.write_bytes(content)

    source = event.source
    conn = db()
    try:
        with conn, conn.cursor() as cur:
            cur.execute(
                """INSERT INTO raw_files
                   (line_message_id, line_group_id, line_user_id, file_type, storage_path, is_slip)
                   VALUES (%s, %s, %s, %s, %s, %s) RETURNING id""",
                (msg.id, getattr(source, "group_id", None), getattr(source, "user_id", None),
                 "image" if is_image else "file", str(local_path), is_image),
            )
            raw_file_id = cur.fetchone()[0]
    finally:
        conn.close()

    log.info("backed up %s -> raw_files.id=%s", local_path, raw_file_id)

    if is_image:
        try:
            extract_and_store(local_path, raw_file_id)
        except Exception:
            log.exception("extract failed for %s", local_path)


def handle_text(event: MessageEvent):
    msg = event.message
    source = event.source
    conn = db()
    try:
        with conn, conn.cursor() as cur:
            cur.execute(
                """INSERT INTO line_messages
                   (line_message_id, line_group_id, line_user_id, message_type, text)
                   VALUES (%s, %s, %s, %s, %s)""",
                (msg.id, getattr(source, "group_id", None), getattr(source, "user_id", None),
                 "text", msg.text),
            )
    finally:
        conn.close()
    log.info("backed up text message %s", msg.id)


def extract_and_store(path: Path, raw_file_id: int):
    ref_qr, _bank_code = decode_qr(str(path))
    s, usage = extract_slip(claude, str(path))
    dirn = direction(s.sender_name + s.sender_account, s.receiver_name + s.receiver_account)
    cost_usd = compute_cost(MODEL, usage.input_tokens, usage.output_tokens)

    conn = db()
    try:
        with conn, conn.cursor() as cur:
            category = guess_category(s.memo, get_categories(cur))
            cur.execute(
                """INSERT INTO slip_transactions
                   (raw_file_id, bank, txn_date, txn_time, amount, fee, sender_name, sender_account,
                    receiver_name, receiver_account, memo, printed_ref, qr_trans_ref, direction, category, ai_model)
                   VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                (raw_file_id, s.bank, s.date, s.time, s.amount, s.fee, s.sender_name, s.sender_account,
                 s.receiver_name, s.receiver_account, s.memo, s.printed_ref, ref_qr, dirn,
                 category, MODEL),
            )
            cur.execute("UPDATE raw_files SET processed = true WHERE id = %s", (raw_file_id,))
            cur.execute(
                """INSERT INTO ai_usage_log (model, input_tokens, output_tokens, cost_usd)
                   VALUES (%s, %s, %s, %s)""",
                (MODEL, usage.input_tokens, usage.output_tokens, cost_usd),
            )
    finally:
        conn.close()

    log.info("slip stored: %s %.2f %s", s.bank, s.amount, s.memo)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
