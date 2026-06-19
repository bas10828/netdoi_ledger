"""POC: read Thai bank-transfer slip images into an income/expense ledger.

Pipeline per image:
  1. Decode the slip-verification QR (opencv) -> transaction reference + sending
     bank code. Used as a dedup key and to cross-check the printed reference.
  2. Send the image to Claude (vision) and extract structured fields as JSON.
  3. Derive direction (expense/income) by comparing the sender against the
     account we own (OWNER_ACCOUNT_HINTS).
  4. Append a row to ledger.csv, skipping references already seen.

Run:
  pip install -r requirements.txt
  copy .env.example to .env and fill in ANTHROPIC_API_KEY=sk-ant-...
  python extract.py
"""

import base64
import csv
import glob
import os
import sys

import anthropic
import cv2
from dotenv import load_dotenv
from pydantic import BaseModel
from pyzbar.pyzbar import decode as zbar_decode

load_dotenv()

# --- config -----------------------------------------------------------------

SLIP_DIR = "slip"
OUT_CSV = "ledger.csv"
MODEL = "claude-sonnet-4-6"

# USD per 1M tokens (input, output)
PRICING_USD_PER_1M = {
    "claude-sonnet-4-6": (3.0, 15.0),
}


def compute_cost(model, input_tokens, output_tokens):
    in_price, out_price = PRICING_USD_PER_1M.get(model, (0.0, 0.0))
    return (input_tokens / 1_000_000) * in_price + (output_tokens / 1_000_000) * out_price

# Substrings identifying the account we own. A slip whose *sender* matches one
# of these is money going OUT (expense); whose *receiver* matches is money
# coming IN (income). Set via OWNER_ACCOUNT_HINTS env var, comma-separated.
OWNER_ACCOUNT_HINTS = [h.strip() for h in os.environ.get("OWNER_ACCOUNT_HINTS", "").split(",") if h.strip()]


# --- structured output schema -----------------------------------------------

class Slip(BaseModel):
    bank: str                  # issuing bank, e.g. "ttb", "Bangkok Bank"
    date: str                  # ISO date YYYY-MM-DD (convert Buddhist year -543)
    time: str                  # HH:MM (24h)
    amount: float              # transfer amount in THB
    fee: float                 # fee in THB (0 if none)
    sender_name: str
    sender_account: str        # may be masked, e.g. "XXX-X-XX123-4"
    receiver_name: str
    receiver_account: str      # account no. or biller/tax id
    memo: str                  # บันทึก / บันทึกช่วยจำ — the expense note
    printed_ref: str           # รหัสอ้างอิง / เลขที่อ้างอิง printed on the slip


EXTRACT_PROMPT = (
    "This is a Thai bank money-transfer slip. Extract the fields into the schema. "
    "Convert the Thai Buddhist-era year to Gregorian (subtract 543; e.g. 69 -> 2026). "
    "Amount and fee are numbers in THB without commas. "
    "memo is the บันทึก / บันทึกช่วยจำ note (the reason for the transfer); empty string if absent. "
    "printed_ref is the รหัสอ้างอิง or เลขที่อ้างอิง value."
)


# --- helpers ----------------------------------------------------------------

def decode_qr(path):
    """Return (transaction_ref, sending_bank_code) from the slip QR, or (None, None).

    Thai slip QR is EMVCo TLV: tag 00 -> sub 01 = bank code, sub 02 = ref.
    """
    img = cv2.imread(path)
    if img is None:
        return None, None
    results = zbar_decode(img)
    if not results:
        ok, infos, _, _ = cv2.QRCodeDetector().detectAndDecodeMulti(img)
        results = [type("D", (), {"data": s.encode()}) for s in (infos if ok else []) if s]
    for d in results:
        s = d.data.decode("utf-8", "replace")
        if not s.startswith("00"):
            continue
        i, bank, ref = 0, None, None
        while i + 4 <= len(s):           # walk top-level TLV
            tag, ln = s[i:i + 2], int(s[i + 2:i + 4])
            val = s[i + 4:i + 4 + ln]
            if tag == "00":              # nested: 00=ver 01=bank 02=ref
                j = 0
                while j + 4 <= len(val):
                    t2, l2 = val[j:j + 2], int(val[j + 2:j + 4])
                    v2 = val[j + 4:j + 4 + l2]
                    if t2 == "01":
                        bank = v2
                    elif t2 == "02":
                        ref = v2
                    j += 4 + l2
            i += 4 + ln
        if ref:
            return ref, bank
    return None, None


def direction(sender, receiver):
    """expense if we are the sender, income if we are the receiver, else unknown."""
    s, r = sender or "", receiver or ""
    if any(h in s for h in OWNER_ACCOUNT_HINTS):
        return "expense"
    if any(h in r for h in OWNER_ACCOUNT_HINTS):
        return "income"
    return "unknown"


def extract_slip(client, path):
    with open(path, "rb") as f:
        b64 = base64.standard_b64encode(f.read()).decode()
    resp = client.messages.parse(
        model=MODEL,
        max_tokens=2000,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {
                    "type": "base64", "media_type": "image/jpeg", "data": b64}},
                {"type": "text", "text": EXTRACT_PROMPT},
            ],
        }],
        output_format=Slip,
    )
    return resp.parsed_output, resp.usage


# --- main -------------------------------------------------------------------

def main():
    if not os.environ.get("ANTHROPIC_API_KEY"):
        sys.exit("set ANTHROPIC_API_KEY first (env var or .env file)")

    client = anthropic.Anthropic()
    files = sorted(glob.glob(os.path.join(SLIP_DIR, "*.jpg")))
    seen = set()
    rows = []

    for path in files:
        ref_qr, bank_code = decode_qr(path)
        if ref_qr and ref_qr in seen:
            print(f"skip duplicate {os.path.basename(path)} (ref {ref_qr})")
            continue
        s, _usage = extract_slip(client, path)
        dedup_key = ref_qr or s.printed_ref
        if dedup_key in seen:
            print(f"skip duplicate {os.path.basename(path)} (ref {dedup_key})")
            continue
        seen.add(dedup_key)
        rows.append({
            "file": os.path.basename(path),
            "date": s.date, "time": s.time,
            "direction": direction(s.sender_name + s.sender_account,
                                   s.receiver_name + s.receiver_account),
            "amount": f"{s.amount:.2f}", "fee": f"{s.fee:.2f}",
            "bank": s.bank,
            "sender": s.sender_name, "receiver": s.receiver_name,
            "memo": s.memo,
            "ref": ref_qr or s.printed_ref, "bank_code": bank_code or "",
        })
        print(f"ok {os.path.basename(path)}: {s.amount:.2f}  {s.memo}")

    with open(OUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)

    exp = sum(float(r["amount"]) for r in rows if r["direction"] == "expense")
    inc = sum(float(r["amount"]) for r in rows if r["direction"] == "income")
    print(f"\n{len(rows)} rows -> {OUT_CSV}")
    print(f"expense {exp:.2f}  income {inc:.2f}  net {inc - exp:.2f}")


if __name__ == "__main__":
    main()
