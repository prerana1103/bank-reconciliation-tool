# Evaluation: engine vs. answer key (synthetic dataset)

- Bank lines: **1130**, ledger lines: **1136** (2266 total)
- Injected discrepancies: **84**
- Detected with correct exception type: **84/84** (100.0%)
- Detected with correct type AND sub-type: **84/84** (100.0%)
- False positives (exceptions on clean transactions): **0**
- Exceptions raised: **84** (reviewer sees 84 items instead of 2266 lines, a 96.3% reduction in lines to inspect)
- Unexplained difference in reconciliation bridge: **0.00**

| Category | Injected | Detected | Correct sub-type |
|---|---:|---:|---:|
| BANK_CHARGE_OR_INTEREST | 10 | 10 | 10 |
| DUP_BANK | 8 | 8 | 8 |
| DUP_LEDGER | 8 | 8 | 8 |
| FEE_VARIANCE | 8 | 8 | 8 |
| IN_TRANSIT | 6 | 6 | 6 |
| LOW_CONFIDENCE | 5 | 5 | 5 |
| ROUNDING | 5 | 5 | 5 |
| SIGN_REVERSAL | 4 | 4 | 4 |
| SPLIT | 6 | 6 | 6 |
| STALE_CHEQUE | 8 | 8 | 8 |
| TRANSPOSITION | 6 | 6 | 6 |
| UNRECORDED_PAYMENT | 6 | 6 | 6 |
| UNRECORDED_RECEIPT | 4 | 4 | 4 |

## Stress test: all reference numbers removed from both files
- Injected issues still classified correctly: **79/84**
- False positives: **294** (exceptions raised: 380)
- Why: without references the engine falls back to amount + date (3-day window) + description, so cheques that clear 4-5 days late stop matching. Reference numbers are the single biggest driver of accuracy.
