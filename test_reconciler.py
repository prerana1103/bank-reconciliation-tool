#!/usr/bin/env python3
"""Run:  python test_reconciler.py   (or: pytest test_reconciler.py)"""
import tempfile
from pathlib import Path

from bank_reconciler import (Config, Reconciler, append_audit_log, load_transactions,
                             parse_amount, verify_audit_log)
from generate_sample_data import EXPECTED, EXPECTED_SUBTYPES, write_csvs


def _run():
    tmp = Path(tempfile.mkdtemp())
    write_csvs(tmp)
    cfg = Config()
    bank, _, _ = load_transactions(tmp / "sample_bank_statement.csv", "BANK", cfg)
    ledger, _, _ = load_transactions(tmp / "sample_ledger.csv", "LEDGER", cfg)
    return Reconciler(bank, ledger, cfg).run()


def test_parse_amount():
    assert parse_amount("1,234.50") == 123450
    assert parse_amount("(99.10)") == -9910
    assert parse_amount("250.00 DR") == -25000
    assert parse_amount("abc") is None and parse_amount("") is None


def test_exception_counts_match_injected_issues():
    rec = _run()
    assert rec.counts_by_type() == EXPECTED, rec.counts_by_type()


def test_subtype_classification():
    subs = {e.subtype for e in _run().exceptions}
    assert EXPECTED_SUBTYPES <= subs, EXPECTED_SUBTYPES - subs


def test_reconciliation_bridge_is_exact():
    assert _run().bridge()["unexplained"] == 0


def test_every_transaction_accounted_for():
    rec = _run()
    for t in rec.bank + rec.ledger:
        assert t.uid in rec.matched or t.uid in rec.txn_exc, t.uid


def test_one_to_one_integrity():
    rec = _run()
    seen = [t.uid for m in rec.matches for t in m.bank + m.ledger]
    assert len(seen) == len(set(seen)), "a transaction was matched twice"


def test_audit_log_detects_tampering():
    p = Path(tempfile.mkdtemp()) / "log.jsonl"
    for i in range(3):
        append_audit_log(p, {"run_id": str(i)})
    assert verify_audit_log(p)[0]
    p.write_text(p.read_text().replace('"run_id": "1"', '"run_id": "X"'))
    assert not verify_audit_log(p)[0]


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print(f"PASS  {name}")
