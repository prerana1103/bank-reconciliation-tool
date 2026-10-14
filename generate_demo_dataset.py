#!/usr/bin/env python3
"""
Generate a quarter (Jul-Sep 2026) of synthetic bank + ledger data (~1,100 lines per side)
with a documented answer key of injected discrepancies, for measuring the engine.

Outputs (demo_dataset/):
    demo_bank_statement.csv   Transaction Date, Narration, Reference, Withdrawal, Deposit, Balance
    demo_ledger.csv           Date, Description, Doc No, Amount
    answer_key.csv            one row per transaction involved in an injected issue
"""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

START = date(2026, 7, 1)
LAST_DAY = 91  # 30 Sep 2026

CUSTOMERS = ["Acme Traders", "Bluefin Logistics", "Crestview Retail", "Delta Foods",
             "Everest Textiles", "Falcon Auto Parts", "Greenleaf Agro", "Horizon Pharma",
             "Indus Packaging", "Jupiter Steel", "Kingsley Chemicals", "Lakeview Hotels"]
VENDORS = ["Kestrel Office Supplies", "Pinnacle Print", "Zenith Hardware", "Orchid Catering",
           "Summit Security", "Vertex Consulting", "Willow Furnishings", "Yonder Travel",
           "Atlas Freight", "Birch Packaging"]
UTILITIES = ["Lumen Power Utility", "Nimbus Cloud Services", "Meridian Couriers",
             "Quartz Cleaning", "Rapid Repairs", "Orion Stationers"]

# category -> (expected exception type, expected sub-type)
CATEGORIES = {
    "DUP_BANK": ("DUPLICATE_BANK", "REPEATED_BANK_ENTRY"),
    "DUP_LEDGER": ("DUPLICATE_LEDGER", "DOUBLE_POSTED_ENTRY"),
    "TRANSPOSITION": ("AMOUNT_MISMATCH", "POSSIBLE_TRANSPOSITION"),
    "FEE_VARIANCE": ("AMOUNT_MISMATCH", "BANK_AMOUNT_LOWER"),
    "ROUNDING": ("AMOUNT_MISMATCH", "ROUNDING_DIFFERENCE"),
    "SIGN_REVERSAL": ("AMOUNT_MISMATCH", "SIGN_REVERSAL"),
    "BANK_CHARGE_OR_INTEREST": ("UNMATCHED_BANK", "UNRECORDED_BANK_CHARGE_OR_INTEREST"),
    "UNRECORDED_PAYMENT": ("UNMATCHED_BANK", "UNRECORDED_PAYMENT"),
    "UNRECORDED_RECEIPT": ("UNMATCHED_BANK", "UNRECORDED_RECEIPT"),
    "STALE_CHEQUE": ("UNMATCHED_LEDGER", "STALE_OUTSTANDING_PAYMENT"),
    "IN_TRANSIT": ("UNMATCHED_LEDGER", "OUTSTANDING_IN_TRANSIT"),
    "SPLIT": ("SPLIT_MATCH", "BATCH_DEPOSIT_OR_PAYMENT"),
    "LOW_CONFIDENCE": ("LOW_CONFIDENCE_MATCH", "NO_REFERENCE_OR_DESCRIPTION_SUPPORT"),
}
COUNTS = {"DUP_BANK": 8, "DUP_LEDGER": 8, "TRANSPOSITION": 6, "FEE_VARIANCE": 8, "ROUNDING": 5,
          "SIGN_REVERSAL": 4, "BANK_CHARGE_OR_INTEREST": 10, "UNRECORDED_PAYMENT": 6,
          "UNRECORDED_RECEIPT": 4, "STALE_CHEQUE": 8, "IN_TRANSIT": 6, "SPLIT": 6,
          "LOW_CONFIDENCE": 5}


def build_demo(seed: int = 7):
    rnd = random.Random(seed)
    used, bank, ledger, seq = set(), [], [], {}

    def amount(lo, hi):
        while True:
            a = round(rnd.uniform(lo, hi), 2)
            if a not in used and a != int(a):
                used.add(a)
                return a

    def day(o):
        return START + timedelta(days=o)

    def issue(cat):
        seq[cat] = seq.get(cat, 0) + 1
        return f"{cat}-{seq[cat]:03d}"

    def led(dt, desc, doc, amt, iid=None):
        row = {"Date": dt.isoformat(), "Description": desc, "Doc No": doc,
               "Amount": f"{amt:.2f}", "_issue": iid}
        ledger.append(row)
        return row

    def bnk(dt, narr, ref, amt, iid=None):
        row = {"Transaction Date": dt.isoformat(), "Narration": narr, "Reference": ref,
               "Withdrawal": f"{-amt:.2f}" if amt < 0 else "",
               "Deposit": f"{amt:.2f}" if amt > 0 else "", "_issue": iid}
        bank.append(row)
        return row

    # ---------------- clean population ----------------
    receipts, payments = [], []
    for i in range(600):
        cust, inv, a = CUSTOMERS[i % len(CUSTOMERS)], 1001 + i, amount(3000, 95000)
        d0, lag = rnd.randint(0, 85), rnd.choice([0, 0, 1, 1, 2])
        lr = led(day(d0), f"Receipt from {cust} - INV-{inv}", f"INV-{inv}", a)
        br = bnk(day(d0 + lag), f"NEFT-{cust.upper()}-INV{inv}",
                 f"INV-{inv}" if rnd.random() > 0.25 else "", a)
        receipts.append((lr, br))
    for i in range(450):
        vend, chq, a = VENDORS[i % len(VENDORS)], 5001 + i, -amount(500, 60000)
        d0, lag = rnd.randint(0, 85), rnd.randint(0, 5)
        lr = led(day(d0), f"Cheque payment to {vend} - CHQ-{chq:06d}", f"CHQ-{chq:06d}", a)
        br = bnk(day(d0 + lag), f"CHQ {chq:06d} {vend.upper()}", f"{chq:06d}", a)
        payments.append((lr, br))
    for m in range(3):
        for k, u in enumerate(UTILITIES):
            a, d0 = -amount(2000, 9000), m * 30 + 3 + k * 4
            led(day(d0), f"{u} - bill", "", a)
            bnk(day(d0 + rnd.randint(0, 2)), f"ACH-{u.upper()}", "", a)

    # ---------------- injected issues ----------------
    # Duplicates (copy appended after original, same date)
    for idx in rnd.sample(range(len(receipts)), COUNTS["DUP_BANK"]):
        iid, (lr, br) = issue("DUP_BANK"), receipts[idx]
        br["_issue"] = iid
        cp = dict(br)
        bank.append(cp)
    for idx in rnd.sample(range(len(payments)), COUNTS["DUP_LEDGER"]):
        iid, (lr, br) = issue("DUP_LEDGER"), payments[idx]
        lr["_issue"] = iid
        ledger.append(dict(lr))

    def swap_digits(a):
        s = f"{a:.2f}".replace(".", "")
        idxs = [i for i in range(len(s) - 3)
                if s[i] != s[i + 1] and not (i == 0 and s[1] == "0")]
        i = rnd.choice(idxs)
        s = s[:i] + s[i + 1] + s[i] + s[i + 2:]
        return int(s) / 100

    for n in range(COUNTS["TRANSPOSITION"]):
        iid, inv, cust = issue("TRANSPOSITION"), 9100 + n, CUSTOMERS[n]
        a, d0 = amount(5000, 90000), rnd.randint(0, 80)
        b = swap_digits(a)
        used.add(b)
        led(day(d0), f"Receipt from {cust} - INV-{inv}", f"INV-{inv}", a, iid)
        bnk(day(d0 + rnd.randint(0, 2)), f"NEFT-{cust.upper()}-INV{inv}", f"INV-{inv}", b, iid)

    for n in range(COUNTS["FEE_VARIANCE"]):
        iid, inv, cust = issue("FEE_VARIANCE"), 9200 + n, CUSTOMERS[n]
        a, d0 = amount(5000, 80000), rnd.randint(0, 80)
        b = round(a - rnd.uniform(8, 180), 2)
        used.add(b)
        led(day(d0), f"Receipt from {cust} - INV-{inv}", f"INV-{inv}", a, iid)
        bnk(day(d0 + rnd.randint(0, 2)), f"WIRE-{cust.upper()}-INV{inv}", f"INV-{inv}", b, iid)

    for n in range(COUNTS["ROUNDING"]):
        iid, inv, cust = issue("ROUNDING"), 9300 + n, CUSTOMERS[n]
        a, d0 = amount(1000, 40000), rnd.randint(0, 80)
        b = round(a + rnd.choice([-1, 1]) * rnd.uniform(0.05, 0.95), 2)
        used.add(b)
        led(day(d0), f"Receipt from {cust} - INV-{inv}", f"INV-{inv}", a, iid)
        bnk(day(d0 + rnd.randint(0, 2)), f"NEFT-{cust.upper()}-INV{inv}", f"INV-{inv}", b, iid)

    for n in range(COUNTS["SIGN_REVERSAL"]):
        iid, chq, vend = issue("SIGN_REVERSAL"), 9001 + n, VENDORS[n]
        a, d0 = amount(1000, 30000), rnd.randint(0, 80)
        led(day(d0), f"Cheque payment to {vend} - CHQ-{chq:06d}", f"CHQ-{chq:06d}", a, iid)
        bnk(day(d0 + rnd.randint(0, 3)), f"CHQ {chq:06d} {vend.upper()}", f"{chq:06d}", -a, iid)

    # Bank-only items
    charge_days = [30, 30, 61, 61, 91, 91, 15, 45, 75, 91]
    charge_desc = [("BANK CHARGES", -1), ("SMS ALERT CHARGES", -1), ("ANNUAL FEE DEBIT CARD", -1),
                   ("INTEREST CREDIT", 1), ("GST ON CHARGES", -1)]
    for n in range(COUNTS["BANK_CHARGE_OR_INTEREST"]):
        desc, sign = charge_desc[n % len(charge_desc)]
        a = sign * (amount(200, 3000) if sign > 0 else amount(50, 900))
        bnk(day(charge_days[n]), f"{desc} {day(charge_days[n]).strftime('%b').upper()}", "", a,
            issue("BANK_CHARGE_OR_INTEREST"))
    pay_desc = ["DIRECT DEBIT INSURANCE PREMIUM", "EMI AUTO DEBIT LOAN", "STANDING INSTRUCTION RENT",
                "DIRECT DEBIT PROFESSIONAL BODY", "EMI AUTO DEBIT EQUIPMENT", "SI MAINTENANCE"]
    for n in range(COUNTS["UNRECORDED_PAYMENT"]):
        bnk(day(rnd.randint(0, 85)), pay_desc[n], "", -amount(2000, 30000), issue("UNRECORDED_PAYMENT"))
    for n in range(COUNTS["UNRECORDED_RECEIPT"]):
        bnk(day(rnd.randint(0, 85)), f"NEFT CR UNKNOWN REMITTER {7000 + n}", "",
            amount(1000, 40000), issue("UNRECORDED_RECEIPT"))

    # Ledger-only items
    for n in range(COUNTS["STALE_CHEQUE"]):
        vend, chq = VENDORS[n], 9500 + n
        led(day(rnd.randint(0, 70)), f"Cheque payment to {vend} - CHQ-{chq:06d}", f"CHQ-{chq:06d}",
            -amount(1500, 40000), issue("STALE_CHEQUE"))
    for n in range(COUNTS["IN_TRANSIT"]):
        cust, inv = CUSTOMERS[n], 9400 + n
        led(day(rnd.randint(88, LAST_DAY)), f"Deposit in transit - {cust} INV-{inv}", f"INV-{inv}",
            amount(3000, 60000), issue("IN_TRANSIT"))

    # Split batch deposits: 1 bank line = 3 ledger receipts
    for n in range(COUNTS["SPLIT"]):
        iid, d0 = issue("SPLIT"), rnd.randint(6, 78)
        parts = [amount(2000, 9000) for _ in range(3)]
        total = round(sum(parts), 2)
        used.add(total)
        for k, p in enumerate(parts, 1):
            led(day(d0 - (k == 1)), f"Cash receipt counter {k} - batch {n + 1}", "", p, iid)
        bnk(day(d0 + rnd.randint(0, 2)), f"CASH DEPOSIT BATCH {n + 1:03d}", "", total, iid)

    # Low confidence: amount + date only
    items = ["Office supplies purchase", "Courier charges", "Printing expense", "Travel expense",
             "Staff welfare expense"]
    for n in range(COUNTS["LOW_CONFIDENCE"]):
        iid, a, d0 = issue("LOW_CONFIDENCE"), -amount(300, 4000), rnd.randint(0, 85)
        led(day(d0), items[n], "", a, iid)
        bnk(day(d0 + rnd.randint(0, 2)), f"POS {90000 + n * 37} MERCHANT", "", a, iid)

    bank.sort(key=lambda r: r["Transaction Date"])
    ledger.sort(key=lambda r: r["Date"])
    return bank, ledger


def write_demo(out_dir: Path):
    bank, ledger = build_demo()
    out_dir.mkdir(parents=True, exist_ok=True)
    key = []
    bal = 500_000_000  # opening balance, cents
    for i, r in enumerate(bank):
        w, dp = r["Withdrawal"], r["Deposit"]
        bal += round(float(dp or 0) * 100) - round(float(w or 0) * 100)
        r["Balance"] = f"{bal / 100:.2f}"
        if r["_issue"]:
            key.append((r["_issue"], "BANK", i + 2))
    for i, r in enumerate(ledger):
        if r["_issue"]:
            key.append((r["_issue"], "LEDGER", i + 2))

    def dump(name, rows, fields):
        with open(out_dir / name, "w", newline="", encoding="utf-8") as f:
            wr = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
            wr.writeheader()
            wr.writerows(rows)

    dump("demo_bank_statement.csv", bank,
         ["Transaction Date", "Narration", "Reference", "Withdrawal", "Deposit", "Balance"])
    dump("demo_ledger.csv", ledger, ["Date", "Description", "Doc No", "Amount"])
    with open(out_dir / "answer_key.csv", "w", newline="", encoding="utf-8") as f:
        wr = csv.writer(f)
        wr.writerow(["issue_id", "category", "expected_type", "expected_subtype", "source", "file_row"])
        for iid, src, row in sorted(key):
            cat = iid.rsplit("-", 1)[0]
            wr.writerow([iid, cat, *CATEGORIES[cat], src, row])
    return len(bank), len(ledger), sum(COUNTS.values())


if __name__ == "__main__":
    nb, nl, ni = write_demo(Path(__file__).parent / "demo_dataset")
    print(f"demo_dataset/: {nb} bank lines, {nl} ledger lines, {ni} injected issues")
