#!/usr/bin/env python3
"""Score the reconciliation engine against demo_dataset/answer_key.csv.

Usage: python evaluate.py [demo_dataset_dir]
"""
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd

from bank_reconciler import Config, Reconciler, load_transactions


def evaluate(folder: Path) -> dict:
    cfg = Config()
    bank, _, _ = load_transactions(folder / "demo_bank_statement.csv", "BANK", cfg)
    ledger, _, _ = load_transactions(folder / "demo_ledger.csv", "LEDGER", cfg)
    rec = Reconciler(bank, ledger, cfg).run()
    key = pd.read_csv(folder / "answer_key.csv", dtype=str)

    uid_of = {("BANK", t.row): t.uid for t in bank}
    uid_of.update({("LEDGER", t.row): t.uid for t in ledger})
    exc_of = {e.exc_id: e for e in rec.exceptions}

    per_cat = defaultdict(lambda: {"injected": 0, "detected": 0, "classified": 0})
    issue_uids = set()
    misses = []
    for iid, g in key.groupby("issue_id"):
        cat, etype, esub = g.iloc[0][["category", "expected_type", "expected_subtype"]]
        uids = {uid_of[(r.source, int(r.file_row))] for r in g.itertuples()}
        issue_uids |= uids
        found = [exc_of[x] for x in {rec.txn_exc[u] for u in uids if u in rec.txn_exc}]
        det = any(e.type == etype for e in found)
        cls = any(e.type == etype and e.subtype == esub for e in found)
        per_cat[cat]["injected"] += 1
        per_cat[cat]["detected"] += det
        per_cat[cat]["classified"] += cls
        if not cls:
            misses.append((iid, etype, esub, [(e.type, e.subtype) for e in found]))

    false_pos = [e for e in rec.exceptions
                 if not any(u.strip() in issue_uids
                            for u in (e.bank_ids + "; " + e.ledger_ids).split(";"))]
    n_inj = sum(v["injected"] for v in per_cat.values())
    n_det = sum(v["detected"] for v in per_cat.values())
    n_cls = sum(v["classified"] for v in per_cat.values())
    total_lines = len(bank) + len(ledger)
    return {"rec": rec, "per_cat": dict(per_cat), "n_inj": n_inj, "n_det": n_det,
            "n_cls": n_cls, "false_pos": false_pos, "misses": misses,
            "bank": len(bank), "ledger": len(ledger), "total_lines": total_lines,
            "exceptions": len(rec.exceptions), "bridge": rec.bridge()["unexplained"]}


def report(r: dict) -> str:
    L = ["# Evaluation: engine vs. answer key (synthetic dataset)", "",
         f"- Bank lines: **{r['bank']}**, ledger lines: **{r['ledger']}** "
         f"({r['total_lines']} total)",
         f"- Injected discrepancies: **{r['n_inj']}**",
         f"- Detected with correct exception type: **{r['n_det']}/{r['n_inj']}** "
         f"({r['n_det'] / r['n_inj']:.1%})",
         f"- Detected with correct type AND sub-type: **{r['n_cls']}/{r['n_inj']}** "
         f"({r['n_cls'] / r['n_inj']:.1%})",
         f"- False positives (exceptions on clean transactions): **{len(r['false_pos'])}**",
         f"- Exceptions raised: **{r['exceptions']}** "
         f"(reviewer sees {r['exceptions']} items instead of {r['total_lines']} lines, "
         f"a {1 - r['exceptions'] / r['total_lines']:.1%} reduction in lines to inspect)",
         f"- Unexplained difference in reconciliation bridge: **{r['bridge'] / 100:,.2f}**", "",
         "| Category | Injected | Detected | Correct sub-type |", "|---|---:|---:|---:|"]
    for cat, v in r["per_cat"].items():
        L.append(f"| {cat} | {v['injected']} | {v['detected']} | {v['classified']} |")
    if r["misses"]:
        L += ["", "## Misses / misclassifications"]
        L += [f"- {m[0]}: expected {m[1]}/{m[2]}, got {m[3]}" for m in r["misses"]]
    if r["false_pos"]:
        L += ["", "## False positives"]
        L += [f"- {e.exc_id} {e.type}/{e.subtype}: {e.description[:80]}" for e in r["false_pos"]]
    return "\n".join(L) + "\n"


def stress_no_references(folder: Path) -> dict:
    """Harder test: strip every reference / invoice / cheque number from both files."""
    import re
    import shutil
    import tempfile
    tmp = Path(tempfile.mkdtemp())
    shutil.copy(folder / "answer_key.csv", tmp / "answer_key.csv")
    clean = lambda s: re.sub(r"\s+", " ", re.sub(r"(INV|CHQ)[-\s]?\d+", "", str(s))).strip(" -")
    b = pd.read_csv(folder / "demo_bank_statement.csv", dtype=str, keep_default_na=False)
    b["Reference"], b["Narration"] = "", b["Narration"].map(clean)
    b.to_csv(tmp / "demo_bank_statement.csv", index=False)
    l = pd.read_csv(folder / "demo_ledger.csv", dtype=str, keep_default_na=False)
    l["Doc No"], l["Description"] = "", l["Description"].map(clean)
    l.to_csv(tmp / "demo_ledger.csv", index=False)
    return evaluate(tmp)


if __name__ == "__main__":
    folder = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "demo_dataset"
    res = evaluate(folder)
    text = report(res)
    st = stress_no_references(folder)
    text += ("\n## Stress test: all reference numbers removed from both files\n"
             f"- Injected issues still classified correctly: **{st['n_cls']}/{st['n_inj']}**\n"
             f"- False positives: **{len(st['false_pos'])}** (exceptions raised: {st['exceptions']})\n"
             "- Why: without references the engine falls back to amount + date (3-day window) + "
             "description, so cheques that clear 4-5 days late stop matching. "
             "Reference numbers are the single biggest driver of accuracy.\n")
    (folder / "evaluation_report.md").write_text(text, encoding="utf-8")
    print(text)
