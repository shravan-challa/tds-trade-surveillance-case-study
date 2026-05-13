# Trade Surveillance Data — Case Study

A defensible, reproducible data-quality and surveillance-fitness audit of a synthetic trade-data feed. Built as a take-home case study for a trade-surveillance data role; the framing is "what would a regulator-grade DQ assessment look like for a vendor surveillance feed?"

> The dataset (`data/synthetic_trade_data.csv`) is **completely synthetic**, seeded by the case-study author to mimic the structure of a real surveillance feed. Defect rates of 1–10% per field were embedded deliberately. No real trading activity, no real counterparties, no real instruments.

## TL;DR

- **P1 (ingestion readiness)**: recovered 242,429 of 250,714 raw lines (96.70% parse rate, zero quarantine) by fixing CRLF/bare-CR mixing and `skipinitialspace=True` for the multi-line quoted cells. The remaining 3.30% are continuation bytes inside multi-line cells, not lost records.
- **P2 (accuracy & efficacy)**: 16 surveillance-domain checks across Lineage / Lifecycle / Instrument / Account / Trader / Signatures / Statistical. **The data is now ingestable but is not yet surveillable** — four headline defects (Instrument↔ISIN drift on 88% of instruments, 65% of fills missing Trader, Account collapsed to CounterPartyFirm on 97% of rows, 10% orphaned parent references) would render the surveillance system's alerts unreliable.
- **P3 (metrics)**: 44 KPIs across 7 themes (M1–M7) with RED/AMBER/GREEN status and named owner. The producer owns the upstream defects; the surveillance team owns the rule logic.
- **Tests**: 63 assertions across two test scripts; if any number drifts from the file, a test breaks.

## Bundle (read in this order)

| File | What's inside |
|------|---------------|
| `deliverable/00_executive_summary.docx` | One-page bottom line + four headline findings + ownership matrix + productization opportunity |
| `deliverable/02_productization_proposal.docx` | Standalone product pitch — four framings for turning this case study into a recurring TD surveillance capability (Vendor Feed Onboarding Service / Producer Trust Scorecard / Surveillance Audit Pack / Surveillance Data SLOs), each with MVP scope, success metrics, owner; plus M1–M7 maturity model and open decisions |
| `deliverable/stage1_profile_audit.docx` | Stage 1 — column-by-column raw-data audit, severity-coded flags, raw-line evidence |
| `deliverable/stage_p1_repair_report.docx` | Stage P1 — every named normalisation rule (N1..N11) with rationale and rows-affected count |
| `deliverable/stage_p2_accuracy_efficacy.docx` | Stage P2 — surveillance-domain checks, A1..G2 |
| `deliverable/stage_p3_metrics_scorecard.docx` | Stage P3 — 44 KPIs across 7 themes, RED/AMBER/GREEN, owner, target |
| `deliverable/methodology_journal.docx` | Chronological narrative — every decision, hypothesis, recalibration, lesson |

## How to reproduce

```bash
# From the repo root
cd analysis

# Stage 1 — raw profile + audit
python 01_profile.py
python 02_extract_examples.py
python 03_generate_stage1_doc.py
python 04_line_count_audit.py

# Stage P1 — structural repair + value normalisation
python 05_repair_structural.py
python 06_normalize_values.py
python 07_post_p1_profile.py
python 08_generate_p1_doc.py

# Stage P2 — surveillance-domain checks
python 09_p2_accuracy_efficacy.py
python 10_generate_p2_doc.py

# Methodology journal
python 11_generate_methodology_journal.py

# Tests — 63 assertions across P1 invariants + P2 finding reproducibility
python 12_test_cleaned_p1.py     # 42 P1 invariant tests
python 13_test_p2_findings.py    # 21 P2 finding reproducibility tests

# Stage P3 — surveillance DQ scorecard
python 14_p3_metrics.py
python 15_generate_p3_doc.py

# Executive summary + standalone productization proposal
python 16_generate_exec_summary.py
python 17_generate_productization_proposal.py
```

Every artifact (TSV, JSON, docx) regenerates without manual intervention. The pipeline outputs three TSVs that are checked in as part of the audit trail:

- `analysis/cleaned_stage_p1.tsv` — output of `05_repair_structural.py` (structural repair only, before value normalisation)
- `analysis/cleaned_p1_final.tsv` — **the P1 deliverable**: 242,429 rows after both structural repair and value normalisation, ready for ingestion
- `analysis/06_row_flags.tsv` — per-row flag bitset showing which P1 rules fired on each row (audit trail / row-level evidence)

## Repository layout

```
.
├── README.md                       # this file
├── LICENSE                         # MIT
├── brief.md                        # the take-home case-study brief (decoded)
├── data/
│   └── synthetic_trade_data.csv    # 52 MB — input file, fully synthetic
├── analysis/
│   ├── 01_profile.py ... 16_generate_exec_summary.py    # numbered pipeline
│   ├── *.json                      # frozen audit outputs (per-stage)
│   └── *_examples.json             # row-level evidence per finding
└── deliverable/
    └── *.docx                      # the six deliverables a reviewer reads
```

## Design principles

1. **Every rule has a stable ID and a stated rationale.** N1..N11 for P1, A1..G2 for P2, M1..M7 themes for P3. Every count is reproducible from JSON, not a notebook cell.
2. **"Defensible" is the bar.** Every fix has a stated rule + rationale, not "I cleaned this up". Conservative defaults — when in doubt, null + flag rather than guess.
3. **Fix the cause, not the symptom.** P1 spent 30 minutes finding the one parser flag (`skipinitialspace=True`) that recovered 92% of malformed rows, instead of writing 100,000 row-by-row repairs.
4. **Surveillance lens, not generic DQ lens.** P2 frames every check as "what surveillance rule does this support, and would a regulator be satisfied?" — not "is this column tidy?".
5. **Producer vs surveillance-team ownership is explicit.** Every P3 metric names the owner so the report routes work to the right team. "Rule fires too much" usually means "the data is broken upstream", not "the rule is wrong".

## Tech stack

- Python 3.14
- pandas 3.0.2, numpy 2.4.4
- python-docx 1.2.0
- No notebook. Pipeline is reproducible from CLI in script-order.

## Author

Shravan Challa · cm.shravan@gmail.com
