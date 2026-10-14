# Automated Bank Reconciliation Tool

Matches bank statement lines against accounting records, auto-classifies every discrepancy, and writes an
audit-ready Excel workbook plus a tamper-evident run log.

**Requirements:** Python 3.9+, `pandas`, `openpyxl` (`pip install pandas openpyxl`)

## Quick start

```bash
python generate_sample_data.py          # creates sample_data/ with 13 known injected issues
python bank_reconciler.py \
    --bank sample_data/sample_bank_statement.csv \
    --ledger sample_data/sample_ledger.csv \
    --out reconciliation_report.xlsx
python test_reconciler.py               # verifies the engine finds every injected issue
python bank_reconciler.py --verify-log audit_log.jsonl   # checks the audit-log hash chain
```

## Input formats

CSV or Excel. Columns are auto-detected (`Date`, `Narration`/`Description`, `Reference`/`Doc No`/`Chq No`,
and either `Amount` or `Debit`/`Credit`). Override with `--bank-cols "date=Posting Date,amount=Amt"`.

Sign convention: **positive = money in, negative = money out.** For separate Debit/Credit columns the tool
applies bank convention (debit = outflow) to the statement and cash-book convention (debit = inflow) to the
ledger. If your file differs, use `--bank-sign -1` or `--ledger-sign -1`. Amounts accept `1,234.50`,
`(99.10)`, `250.00 DR`. Rows with an unparseable date or amount are listed on `Audit_Trail`, never dropped silently.

## Matching waterfall

| Rule | Logic | Result |
|---|---|---|
| R1 | Same reference + amount, date within 15 days | clean match |
| R2 | Same amount, date within 3 days, description similarity >= 0.70 | clean match |
| R3 | Same reference, different amount | `AMOUNT_MISMATCH` |
| R4 | Same amount + date within 3 days, nothing else corroborates | `LOW_CONFIDENCE_MATCH` |
| R5 | One bank line = sum of 2-3 ledger entries (or reverse) | `SPLIT_MATCH` |
| R6 | Similar description + date, near amount or sign flipped | `AMOUNT_MISMATCH` |

Matching is strictly one-to-one (a transaction can never be used twice). Money is held as integer cents.

## Exception classification

| Type | Sub-types |
|---|---|
| `UNMATCHED_BANK` | `UNRECORDED_BANK_CHARGE_OR_INTEREST`, `UNRECORDED_RECEIPT`, `UNRECORDED_PAYMENT` |
| `UNMATCHED_LEDGER` | `OUTSTANDING_IN_TRANSIT`, `STALE_OUTSTANDING_PAYMENT`, `AGED_UNCLEARED_RECEIPT` |
| `DUPLICATE_BANK` / `DUPLICATE_LEDGER` | surplus copy of an identical entry in the same source |
| `AMOUNT_MISMATCH` | `SIGN_REVERSAL`, `POSSIBLE_TRANSPOSITION`, `ROUNDING_DIFFERENCE`, `BANK_AMOUNT_LOWER`, `BANK_AMOUNT_HIGHER` |
| `LOW_CONFIDENCE_MATCH`, `SPLIT_MATCH` | paired, but flagged for human confirmation |

Severity (High / Medium / Low) comes from materiality thresholds (`--high-materiality`, default 10,000;
`--medium-materiality`, default 1,000), with floors for duplicates, sign reversals and stale items.
Every exception carries an auto-written explanation and a suggested action.

## Workbook sheets

| Sheet | Purpose |
|---|---|
| `Summary` | Population, match rates, exceptions by type/severity, **reconciliation bridge** (net difference must be fully explained, else FAIL), sign-off status, signature block |
| `Exceptions` | One row per exception, sorted by severity. Yellow cells = reviewer disposition (dropdown), comments, name, date |
| `Matched` | Every pairing with rule, confidence score, date variance, difference |
| `Bank_Txns` / `Ledger_Txns` | Full population with status and links to Match ID / Exception ID |
| `Audit_Trail` | Run ID, user, SHA-256 of both inputs, all parameters, rule counts, rejected rows |

Totals on `Summary` are live Excel formulas. Sign-off status reads **READY FOR SIGN-OFF** only when the
bridge is zero *and* every exception has a disposition.

## Audit trail

Each run appends one line to `audit_log.jsonl` (beside the report): inputs + hashes, parameters, result
counts, and the SHA-256 of the report. Each entry embeds the hash of the previous one, so editing or deleting
a past entry is detected by `--verify-log`. Note: the report hash is taken at export, so re-saving the
workbook in Excel afterwards will (correctly) no longer match it. Keep the original as the evidence copy.

## Tuning

Every threshold is a CLI flag (e.g. `--date-tolerance-days 5`, `--amount-tolerance 0.05`, `--dayfirst`)
or a key in a JSON file passed with `--config`. All values used are recorded on `Audit_Trail`.

## Known limits

- Single account / single currency per run.
- Descriptions are matched by text similarity, so unusual bank narration formats may need `--desc-match-threshold` tuning.
- Split matching considers up to 3 items (`--split-max-items`).
- European number format (`1.234,50`) is not auto-detected.
- Treat output as audit *support*: a reviewer must still disposition each exception.
