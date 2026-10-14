#!/usr/bin/env python3
"""Generate a realistic month of bank + ledger data with KNOWN injected discrepancies."""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

CUSTOMERS = ["Acme Traders", "Bluefin Logistics", "Crestview Retail", "Delta Foods",
             "Everest Textiles", "Falcon Auto Parts", "Greenleaf Agro", "Horizon Pharma",
             "Indus Packaging", "Jupiter Steel"]
VENDORS = ["Kestrel Office Supplies", "Pinnacle Print", "Zenith Hardware", "Orchid Catering",
           "Summit Security", "Vertex Consulting", "Willow Furnishings", "Yonder Travel"]
UTILITIES = ["Lumen Power Utility", "Nimbus Cloud Services", "Meridian Couriers",
             "Quartz Cleaning", "Rapid Repairs", "Orion Stationers"]

# What the engine should find in the generated data
EXPECTED = {
    "UNMATCHED_BANK": 3, "UNMATCHED_LEDGER": 2, "DUPLICATE_BANK": 1, "DUPLICATE_LEDGER": 1,
    "AMOUNT_MISMATCH": 4, "SPLIT_MATCH": 1, "LOW_CONFIDENCE_MATCH": 1,
}
EXPECTED_SUBTYPES = {"SIGN_REVERSAL", "ROUNDING_DIFFERENCE", "POSSIBLE_TRANSPOSITION",
                     "BANK_AMOUNT_LOWER", "OUTSTANDING_IN_TRANSIT", "STALE_OUTSTANDING_PAYMENT"}


def build_sample(seed: int = 2026):
    rnd = random.Random(seed)
    base = date(2026, 9, 1)
    used, bank, ledger = set(), [], []

    def amount(lo, hi):
        while True:
            a = round(rnd.uniform(lo, hi), 2)
            if a not in used and a != int(a):
                used.add(a)
                return a

    def d(offset):
        return base + timedelta(days=offset)

    def led(dt, desc, doc, amt):
        ledger.append({"Date": dt.isoformat(), "Description": desc, "Doc No": doc, "Amount": f"{amt:.2f}"})

    def bnk(dt, narr, ref, amt):
        bank.append({"Transaction Date": dt.isoformat(), "Narration": narr, "Reference": ref,
                     "Withdrawal": f"{-amt:.2f}" if amt < 0 else "",
                     "Deposit": f"{amt:.2f}" if amt > 0 else ""})

    # --- clean customer receipts (reference on both sides, bank lag 0-2 days) ---
    receipts = []
    for i in range(40):
        cust, inv, a = CUSTOMERS[i % len(CUSTOMERS)], 1001 + i, amount(5000, 95000)
        day, lag = rnd.randint(0, 27), rnd.choice([0, 0, 1, 1, 2])
        led(d(day), f"Receipt from {cust} - INV-{inv}", f"INV-{inv}", a)
        bnk(d(day + lag), f"NEFT-{cust.upper()}-INV{inv}", f"INV-{inv}" if rnd.random() > 0.25 else "", a)
        receipts.append((day, inv, cust, a, lag))

    # --- clean cheque payments (bank clears up to 5 days later) ---
    payments = []
    for i in range(30):
        vend, chq, a = VENDORS[i % len(VENDORS)], 5001 + i, -amount(800, 60000)
        day, lag = rnd.randint(0, 25), rnd.randint(0, 5)
        led(d(day), f"Cheque payment to {vend} - CHQ-{chq:06d}", f"CHQ-{chq:06d}", a)
        bnk(d(day + lag), f"CHQ {chq:06d} {vend.upper()}", f"{chq:06d}", a)
        payments.append((day, chq, vend, a, lag))

    # --- recurring utilities: no reference anywhere (exercises description matching) ---
    for i, u in enumerate(UTILITIES):
        a, day = -amount(2000, 9000), 3 + i * 4
        led(d(day), f"{u} - Sep bill", "", a)
        bnk(d(day + rnd.randint(0, 2)), f"ACH-{u.upper()}", "", a)

    # ===================== injected discrepancies ===========================
    # 1. Duplicate bank credit (bank posted the same receipt twice)
    day, inv, cust, a, lag = receipts[3]
    bnk(d(day + lag), f"NEFT-{cust.upper()}-INV{inv}", f"INV-{inv}", a)
    # 2. Duplicate ledger posting (payment entered twice)
    day, chq, vend, a, lag = payments[4]
    led(d(day), f"Cheque payment to {vend} - CHQ-{chq:06d}", f"CHQ-{chq:06d}", a)
    # 3. Transposition: ledger 12,450.00 vs bank 12,540.00
    led(d(10), "Receipt from Acme Traders - INV-9001", "INV-9001", 12450.00)
    bnk(d(11), "NEFT-ACME TRADERS-INV9001", "INV-9001", 12540.00)
    # 4. Bank fee deducted from a wire: ledger 10,000.00 vs bank 9,985.00
    led(d(12), "Receipt from Delta Foods - INV-9002", "INV-9002", 10000.00)
    bnk(d(12), "WIRE-DELTA FOODS-INV9002", "INV-9002", 9985.00)
    # 5. Rounding: ledger 4,310.50 vs bank 4,311.00
    led(d(15), "Receipt from Falcon Auto Parts - INV-9003", "INV-9003", 4310.50)
    bnk(d(16), "NEFT-FALCON AUTO PARTS-INV9003", "INV-9003", 4311.00)
    # 6. Sign reversal: payment posted in the ledger as a receipt
    led(d(8), "Cheque payment to Zenith Hardware - CHQ-009004", "CHQ-009004", 3200.00)
    bnk(d(10), "CHQ 009004 ZENITH HARDWARE", "009004", -3200.00)
    # 7. Unmatched bank items: charges, interest, unrecorded direct debit
    bnk(d(29), "BANK CHARGES SEP", "", -450.00)
    bnk(d(29), "INTEREST CREDIT SEP", "", 1230.55)
    bnk(d(11), "DIRECT DEBIT INSURANCE PREMIUM", "", -7800.00)
    # 8. Unmatched ledger items: stale cheque + deposit in transit
    led(d(2), "Cheque payment to Pinnacle Print - CHQ-009005", "CHQ-009005", -20000.00)
    led(d(28), "Deposit in transit - Delta Foods INV-9006", "INV-9006", 15000.00)
    # 9. Split: one bank batch deposit = three ledger counter receipts
    for k, amt in enumerate((10500.00, 8250.00, 6250.00), 1):
        led(d(15 + (k > 1)), f"Cash receipt counter {k}", "", amt)
    bnk(d(17), "CASH DEPOSIT BATCH 0917", "", 25000.00)
    # 10. Low confidence: amount + date only, garbled bank narration
    led(d(12), "Office supplies purchase", "", -842.17)
    bnk(d(13), "POS 99812 MERCHANT", "", -842.17)

    bank.sort(key=lambda r: r["Transaction Date"])
    ledger.sort(key=lambda r: r["Date"])
    return bank, ledger, EXPECTED


def write_csvs(out_dir: Path):
    bank, ledger, _ = build_sample()
    out_dir.mkdir(parents=True, exist_ok=True)
    for name, rows in (("sample_bank_statement.csv", bank), ("sample_ledger.csv", ledger)):
        with open(out_dir / name, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)
    return len(bank), len(ledger)


if __name__ == "__main__":
    nb, nl = write_csvs(Path(__file__).parent / "sample_data")
    print(f"Wrote sample_data/ with {nb} bank rows and {nl} ledger rows")
