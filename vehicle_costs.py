"""Per-vehicle fuel cost analysis from the vehicle_fuel_fills table.

The table is filled by hand from tax-invoice and odometer photos (one row per fill);
this module turns it into what the /vehicles page shows: per-vehicle efficiency,
monthly spend, how much of the spend is down to pump-price changes, and rows that
need a human to check them.
"""

from collections import defaultdict
from datetime import date


# Fills before this date set the "normal" price each fuel type is compared against.
PRICE_BASELINE_END = "2026-03-01"

# Above this, a km/L figure means fills are missing from the log rather than a frugal engine.
PLAUSIBLE_MAX_KM_PER_L = {"diesel": 16.0, "benzine": 13.0}
FUEL_LABEL = {"diesel": "ดีเซล", "benzine": "เบนซิน"}

# Note fragments that mean a row needs a human to look at it, and how to describe it.
ISSUE_RULES = [
    ("ANOTHER COMPANY", "ใบกำกับภาษีออกในชื่อบริษัทอื่น"),
    ("ANOMALY", "ชนิดน้ำมันไม่ตรงกับรถ"),
    ("WRONG PLATE", "ทะเบียนรถบนใบกำกับภาษีผิด"),
    ("plate printed", "ทะเบียนรถบนใบกำกับภาษีพิมพ์ผิด"),
    ("no plate", "ใบเสร็จไม่มีทะเบียนรถ"),
    ("verify", "ต้องตรวจว่าไม่ใช่การเบิกซ้ำ"),
    ("bank slip only", "มีสลิปโอนแต่ไม่มีใบกำกับภาษี"),
    ("mislabeled", "ลงชื่อรถผิดในบัญชี"),
]

THAI_MONTHS = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]


def _num(v):
    return float(v) if v not in (None, "") else None


def load_fills(cur):
    cur.execute(
        """SELECT fill_date, vehicle, plate, fuel_type, fuel_product, price_per_l, liters, amount,
                  odometer_km, note, station
           FROM vehicle_fuel_fills
           ORDER BY fill_date, vehicle, id"""
    )
    fills = [
        {
            "date": d.isoformat(), "vehicle": vehicle, "plate": plate, "fuel_type": fuel_type, "product": product,
            "price": _num(price), "liters": _num(liters), "amount": _num(amount) or 0.0, "odometer": _num(odo),
            "note": note, "station": station,
        }
        for d, vehicle, plate, fuel_type, product, price, liters, amount, odo, note, station in cur.fetchall()
    ]
    _estimate_missing_liters(fills)
    return fills


def _estimate_missing_liters(fills):
    """Rows with only a bank slip have no liters; price them at the nearest receipt of the same fuel."""
    priced = defaultdict(list)
    for f in fills:
        if f["price"]:
            priced[f["fuel_type"]].append((date.fromisoformat(f["date"]), f["price"]))
    for f in fills:
        f["estimated"] = f["liters"] is None
        if f["estimated"] and priced[f["fuel_type"]]:
            d = date.fromisoformat(f["date"])
            f["price"] = min(priced[f["fuel_type"]], key=lambda p: abs((p[0] - d).days))[1]
            f["liters"] = f["amount"] / f["price"]


def _month_label(key):
    y, m = key.split("-")
    return f"{THAI_MONTHS[int(m) - 1]} {y[2:]}"


def _months_between(d0, d1):
    return max((date.fromisoformat(d1) - date.fromisoformat(d0)).days / 30.44, 1e-9)


def _dense_window_km_per_l(rows):
    """km/L over the longest run of consecutive fills that all carry an odometer reading —
    the most trustworthy figure when the overall log has gaps."""
    best, run = None, []
    for f in rows + [None]:
        if f is not None and f["odometer"]:
            run.append(f)
            continue
        if len(run) >= 5:
            km = run[-1]["odometer"] - run[0]["odometer"]
            liters = sum(x["liters"] for x in run[1:])
            if km > 0 and liters and (best is None or len(run) > best[0]):
                best = (len(run), km / liters, run[0]["date"], run[-1]["date"])
        run = []
    return best


def vehicle_summary(fills):
    by_vehicle = defaultdict(list)
    for f in fills:
        by_vehicle[f["vehicle"]].append(f)
    latest_price = {}
    for f in fills:
        if f["price"] and not f["estimated"]:
            latest_price[f["fuel_type"]] = f["price"]

    out = []
    for name, rows in by_vehicle.items():
        fuel = rows[0]["fuel_type"]
        odo = [f for f in rows if f["odometer"]]
        v = {
            "name": name, "plate": rows[0]["plate"], "fuel_type": fuel, "fuel_label": FUEL_LABEL.get(fuel, fuel),
            "fills": len(rows), "amount": sum(f["amount"] for f in rows), "liters": sum(f["liters"] or 0 for f in rows),
            "first_date": rows[0]["date"], "last_date": rows[-1]["date"],
            "km": None, "km_per_l": None, "km_per_month": None, "baht_per_km": None,
            "efficiency_note": "", "missing_liters": None, "odometer_last": odo[-1]["odometer"] if odo else None,
        }
        if len(odo) >= 2:
            a, b = odo[0], odo[-1]
            km = b["odometer"] - a["odometer"]
            # Fuel burned between two readings is what was put in after the first, up to the second.
            ia, ib = rows.index(a), rows.index(b)
            liters = sum(f["liters"] or 0 for f in rows[ia + 1:ib + 1])
            v["km"] = km
            v["km_per_month"] = km / _months_between(a["date"], b["date"])
            if liters:
                kpl = km / liters
                ceiling = PLAUSIBLE_MAX_KM_PER_L.get(fuel, 99)
                dense = _dense_window_km_per_l(rows)
                if kpl > ceiling and dense:
                    v["missing_liters"] = km / dense[1] - liters
                    v["efficiency_note"] = (
                        f"ใบเสร็จไม่ครบ: ถ้าคิดจากใบเสร็จได้ {kpl:.1f} กม./ลิตร ซึ่งสูงเกินจริง — "
                        f"ใช้ค่าช่วงที่มีเลขไมล์ครบ ({dense[2]} ถึง {dense[3]}) แทน"
                    )
                    kpl = dense[1]
                elif kpl > ceiling:
                    v["efficiency_note"] = f"{kpl:.1f} กม./ลิตร สูงเกินจริง — น่าจะมีการเติมที่ไม่ได้บันทึก"
                v["km_per_l"] = kpl
                if latest_price.get(fuel):
                    v["baht_per_km"] = latest_price[fuel] / kpl
        v["issues"] = sum(1 for f in rows if issue_of(f))
        out.append(v)
    out.sort(key=lambda v: -v["amount"])
    return out, latest_price


def issue_of(f):
    note = f["note"] or ""
    for frag, label in ISSUE_RULES:
        if frag.lower() in note.lower():
            return label
    return ""


def monthly(fills, vehicles):
    """Spend per month per vehicle, plus the average pump price per fuel type and how much
    of each month's spend comes from prices above the pre-PRICE_BASELINE_END average."""
    base = {}
    for fuel in {f["fuel_type"] for f in fills}:
        pre = [f for f in fills if f["fuel_type"] == fuel and f["date"] < PRICE_BASELINE_END and not f["estimated"]]
        liters = sum(f["liters"] for f in pre)
        base[fuel] = sum(f["amount"] for f in pre) / liters if liters else None

    months = sorted({f["date"][:7] for f in fills})
    names = [v["name"] for v in vehicles]
    rows = []
    for m in months:
        mf = [f for f in fills if f["date"][:7] == m]
        diesel = [f for f in mf if f["fuel_type"] == "diesel" and not f["estimated"]]
        d_liters = sum(f["liters"] for f in diesel)
        price_effect = sum(f["liters"] * (f["price"] - base[f["fuel_type"]])
                           for f in mf if base.get(f["fuel_type"]) and f["price"])
        rows.append({
            "key": m, "label": _month_label(m), "fills": len(mf),
            "amount": sum(f["amount"] for f in mf), "liters": sum(f["liters"] or 0 for f in mf),
            "by_vehicle": {n: sum(f["amount"] for f in mf if f["vehicle"] == n) for n in names},
            "liters_by_vehicle": {n: sum(f["liters"] or 0 for f in mf if f["vehicle"] == n) for n in names},
            "diesel_price": sum(f["amount"] for f in diesel) / d_liters if d_liters else None,
            "price_effect": price_effect,
        })
    return rows, base


BEFORE_MONTHS = 6


def _shift_month(key, delta):
    y, m = map(int, key.split("-"))
    idx = y * 12 + m - 1 + delta
    return f"{idx // 12:04d}-{idx % 12 + 1:02d}"


def usage_vs_price(fills, vehicles):
    """Did fuel use go up, down or stay flat when prices jumped? Compares average liters per month
    in the BEFORE_MONTHS before PRICE_BASELINE_END with every month since, per vehicle, and splits
    the change in monthly spend into a volume part (more/fewer liters at the old price) and a price
    part (the new liters at the price difference)."""
    start = PRICE_BASELINE_END[:7]
    before_keys = [_shift_month(start, -i) for i in range(BEFORE_MONTHS, 0, -1)]
    last = max(f["date"][:7] for f in fills)
    after_keys, k = [], start
    while k <= last:
        after_keys.append(k)
        k = _shift_month(k, 1)

    def period(rows, keys):
        sel = [f for f in rows if f["date"][:7] in keys]
        liters = sum(f["liters"] or 0 for f in sel)
        amount = sum(f["amount"] for f in sel)
        n = len(keys)
        return {"liters": liters / n, "amount": amount / n, "price": amount / liters if liters else None}

    def compare(rows):
        b, a = period(rows, before_keys), period(rows, after_keys)
        change = (a["liters"] - b["liters"]) / b["liters"] * 100 if b["liters"] else None
        volume = (a["liters"] - b["liters"]) * (b["price"] or 0)
        price = a["liters"] * ((a["price"] or 0) - (b["price"] or 0))
        if change is None:
            verdict = "ไม่มีข้อมูลช่วงก่อน"
        elif abs(change) < 5:
            verdict = "ใช้เท่าเดิม"
        else:
            verdict = "ใช้เพิ่มขึ้น" if change > 0 else "ใช้ลดลง"
        return {"before": b, "after": a, "liters_change": change, "verdict": verdict,
                "spend_change": a["amount"] - b["amount"], "volume_effect": volume, "price_effect": price}

    per_vehicle = []
    for v in vehicles:
        rows = [f for f in fills if f["vehicle"] == v["name"]]
        per_vehicle.append({"name": v["name"], "missing_liters": v["missing_liters"], **compare(rows)})
    # A vehicle with missing receipts makes its own before/after comparison unreliable, so the
    # headline answer uses only vehicles whose log is complete.
    complete = [v["name"] for v in vehicles if not v["missing_liters"]]
    return {
        "complete_names": complete,
        "total_complete": compare([f for f in fills if f["vehicle"] in complete]),
        "before_label": f"{_month_label(before_keys[0])} – {_month_label(before_keys[-1])}",
        "after_label": f"{_month_label(after_keys[0])} – {_month_label(after_keys[-1])}",
        "before_months": len(before_keys), "after_months": len(after_keys),
        "total": compare(fills), "vehicles": per_vehicle,
    }


def build_page_data(cur):
    fills = load_fills(cur)
    if not fills:
        return None
    vehicles, latest_price = vehicle_summary(fills)
    months, base = monthly(fills, vehicles)
    crisis = [m for m in months if m["key"] >= PRICE_BASELINE_END[:7]]
    issues = [{**f, "issue": issue_of(f)} for f in fills if issue_of(f)]
    recent = [m for m in months[-3:]]
    return {
        "fills": fills, "vehicles": vehicles, "months": months, "issues": issues,
        "usage": usage_vs_price(fills, vehicles),
        "baseline": base, "latest_price": latest_price,
        "baseline_label": _month_label(PRICE_BASELINE_END[:7]),
        "kpi": {
            "amount": sum(f["amount"] for f in fills),
            "liters": sum(f["liters"] or 0 for f in fills),
            "fills": len(fills),
            "first_date": fills[0]["date"], "last_date": fills[-1]["date"],
            "price_effect_since": sum(m["price_effect"] for m in crisis),
            "monthly_avg_recent": sum(m["amount"] for m in recent) / len(recent) if recent else 0,
            "monthly_liters_recent": sum(m["liters"] for m in recent) / len(recent) if recent else 0,
            "missing_liters": sum(v["missing_liters"] or 0 for v in vehicles),
        },
    }


# ---------------------------------------------------------------------------
# Bank slips vs receipts
# ---------------------------------------------------------------------------

# Memo keywords that name a vehicle when the slip's category doesn't.
VEHICLE_KEYWORDS = [("vigo", "Vigo"), ("revo", "Revo"), ("d-max", "Dmax"), ("dmax", "Dmax"), ("รถตู้", "รถตู้")]
# A slip paid straight to a station is the fill itself, so it must fall on the receipt's date
# (±1 day for late-night fills). A transfer to a person reimburses a fill made up to this
# many days earlier.
REIMBURSE_WINDOW_DAYS = 7


def _is_station(slip):
    return slip["receiver"].upper().startswith(("PTTST", "PTT ", "PT ", "BANGCHAK", "SHELL", "CALTEX"))


STATUS = {
    "matched": "ตรงกัน",
    "vehicle_mismatch": "ลงชื่อรถในบัญชีไม่ตรงกับใบเสร็จ",
    "slip_only": "มีสลิปโอน ไม่มีใบกำกับภาษี",
    "receipt_only": "มีใบเสร็จ ไม่พบสลิปโอน",
    "before_ledger": "ก่อนเริ่มบันทึกสลิป",
}


def fetch_fuel_slips(cur):
    """Fuel purchases recorded from bank slips (engine oil and other non-fuel 'น้ำมัน' excluded)."""
    cur.execute(
        """SELECT id, txn_date, amount, category, receiver_name, memo
           FROM slip_transactions
           WHERE direction = 'expense'
             AND (category LIKE 'ค่าน้ำมัน%%' OR memo ILIKE '%%เติมน้ำมัน%%' OR memo ILIKE '%%ค่าน้ำมัน%%')
             AND COALESCE(memo, '') NOT ILIKE '%%น้ำมันเครื่อง%%'
           ORDER BY txn_date"""
    )
    slips = []
    for sid, d, amount, category, receiver, memo in cur.fetchall():
        vehicle = None
        if category and " - " in category:
            vehicle = category.split(" - ", 1)[1].strip()
        if not vehicle:
            text = (memo or "").lower()
            vehicle = next((v for kw, v in VEHICLE_KEYWORDS if kw in text), None)
        slips.append({"id": sid, "date": d.isoformat(), "amount": float(amount), "category": category or "",
                      "receiver": receiver or "", "memo": memo or "", "vehicle": vehicle})
    return slips


def reconcile(fills, slips):
    """Line bank slips up with logged fills on one timeline. Station slips are matched first
    (same amount, same day), then reimbursements to people (same amount, fill on or up to
    REIMBURSE_WINDOW_DAYS before the transfer); within those, the same vehicle wins, then the
    nearest date. A fill logged only from a bank slip (no receipt) means the slip has no receipt."""
    ledger_start = min((s["date"] for s in slips), default=None)
    claimed = set()
    rows = []

    def window_ok(slip, fill):
        gap = (_d(slip["date"]) - _d(fill["date"])).days
        return abs(gap) <= 1 if _is_station(slip) else 0 <= gap <= REIMBURSE_WINDOW_DAYS

    for s in sorted(slips, key=lambda s: (not _is_station(s), s["date"])):
        sd = _d(s["date"])
        cands = [(i, f) for i, f in enumerate(fills)
                 if i not in claimed and abs(f["amount"] - s["amount"]) < 0.5 and window_ok(s, f)]
        cands.sort(key=lambda c: (bool(s["vehicle"]) and c[1]["vehicle"] != s["vehicle"],
                                  abs((_d(c[1]["date"]) - sd).days)))
        if not cands:
            rows.append({"date": s["date"], "vehicle": s["vehicle"] or "ไม่ระบุรถ", "amount": s["amount"],
                         "fill": None, "slip": s, "status": "slip_only"})
            continue
        i, f = cands[0]
        claimed.add(i)
        if f["estimated"]:
            status = "slip_only"
        elif s["vehicle"] and s["vehicle"] != f["vehicle"]:
            status = "vehicle_mismatch"
        else:
            status = "matched"
        rows.append({"date": f["date"], "vehicle": f["vehicle"], "amount": f["amount"], "fill": f, "slip": s,
                     "status": status})
    for i, f in enumerate(fills):
        if i in claimed:
            continue
        status = "receipt_only" if ledger_start and f["date"] >= ledger_start else "before_ledger"
        rows.append({"date": f["date"], "vehicle": f["vehicle"], "amount": f["amount"], "fill": f, "slip": None,
                     "status": status})
    rows.sort(key=lambda r: r["date"], reverse=True)
    for r in rows:
        r["status_label"] = STATUS[r["status"]]
    counts = {k: sum(1 for r in rows if r["status"] == k) for k in STATUS}
    gap = {k: sum(r["amount"] for r in rows if r["status"] == k) for k in STATUS}
    return {"rows": rows, "counts": counts, "amounts": gap, "ledger_start": ledger_start}


def _d(s):
    return date.fromisoformat(s)
