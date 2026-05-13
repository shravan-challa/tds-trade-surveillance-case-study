"""
Stage P3 — Surveillance Data Quality Scorecard.

Computes the recurring KPIs a TD surveillance data team would run on every
new vendor / producer feed to decide:
  (a) Can we INGEST this feed?           (Theme M1)
  (b) Are FIELDS we depend on populated? (Theme M2)
  (c) Is the order LINEAGE intact?       (Theme M3)
  (d) Is the LIFECYCLE coherent?         (Theme M4)
  (e) Is the REFERENCE DATA stable?      (Theme M5)
  (f) Can we ATTRIBUTE every execution?  (Theme M6)
  (g) Which RULES can we actually run?   (Theme M7)

Each metric carries:
  id          — stable (M1.1, M2.7, ...)
  name        — one-line label
  definition  — formula in words
  value       — current measurement on cleaned_p1_final.tsv
  unit        — records / pct / ratio / count
  target      — defensible aspiration ("what good looks like")
  status      — GREEN / AMBER / RED vs target
  owner       — Producer / Surveillance team / Both
  rationale   — why this metric matters for surveillance

The scorecard is designed to be re-run on every new feed. The targets are
deliberately conservative — RED means a surveillance rule will be unreliable
or a regulator exam will surface a gap; AMBER means the data is workable but
worth tracking; GREEN means no surveillance impact.

Inputs:
  cleaned_p1_final.tsv                       — P1 output
  06_normalize_stats.json, 06_row_flags.tsv  — P1 audit trail
  07_post_p1_profile.json                    — post-P1 column profile
  09_p2_findings.json                        — P2 surveillance findings

Outputs:
  14_p3_metrics.json                         — full scorecard, ready for the docx renderer
"""

from __future__ import annotations

import json
from pathlib import Path
from datetime import date

import pandas as pd

HERE = Path(__file__).resolve().parent
SRC = HERE / "cleaned_p1_final.tsv"
POST_PROFILE = HERE / "07_post_p1_profile.json"
P2_FINDINGS = HERE / "09_p2_findings.json"
NORMALIZE_STATS = HERE / "06_normalize_stats.json"
OUT = HERE / "14_p3_metrics.json"

CHILD_TYPES = {"Cancel Order", "Replace Order", "Fill", "Reject"}


def status_for(value: float, target_max: float | None = None, amber_max: float | None = None,
               target_min: float | None = None, amber_min: float | None = None) -> str:
    """RED / AMBER / GREEN against either an upper bound (lower is better) or lower bound (higher is better)."""
    if target_max is not None:
        if value <= target_max:
            return "GREEN"
        if amber_max is not None and value <= amber_max:
            return "AMBER"
        return "RED"
    if target_min is not None:
        if value >= target_min:
            return "GREEN"
        if amber_min is not None and value >= amber_min:
            return "AMBER"
        return "RED"
    return "GREEN"


def metric(mid: str, name: str, definition: str, value, unit: str,
           target: str, status: str, owner: str, rationale: str) -> dict:
    return {
        "id": mid,
        "name": name,
        "definition": definition,
        "value": value,
        "unit": unit,
        "target": target,
        "status": status,
        "owner": owner,
        "rationale": rationale,
    }


def main() -> None:
    print(f"Loading {SRC.name} ...")
    df = pd.read_csv(SRC, sep="\t", dtype=str, keep_default_na=True, na_values=[""])
    profile = json.loads(POST_PROFILE.read_text(encoding="utf-8"))
    p2 = json.loads(P2_FINDINGS.read_text(encoding="utf-8"))

    n = len(df)
    raw_lines = profile["stage1_raw_data_lines"]
    print(f"Loaded {n:,} rows from cleaned P1 (raw file had {raw_lines:,} data lines)")

    out: dict = {
        "generated_at": date.today().isoformat(),
        "input_records": n,
        "raw_data_lines": raw_lines,
        "themes": {},
    }

    # ============================================================
    # M1 — Ingestion completeness
    # ============================================================
    parsed_pct = n / raw_lines * 100
    bare_cr = 4193  # from 04_line_count_audit.py — frozen
    crlf = 246522
    quarantined = profile.get("p1_quarantine_rows", 0)

    m1 = [
        metric("M1.1", "Raw data line count",
               "Universal-newline line count of the source CSV minus the header row.",
               raw_lines, "records",
               "n/a (input volume)", "GREEN", "Producer",
               "Anchor for every downstream completeness ratio. Without an authoritative raw count, no ingestion KPI is defensible."),
        metric("M1.2", "Records ingested after P1 repair",
               "Rows in cleaned_p1_final.tsv after structural repair + value normalisation.",
               n, "records",
               f"{int(raw_lines * 0.95):,} (>=95% of raw lines)", "GREEN", "Both",
               "Records the surveillance system actually sees. Direct analogue to the JPM-case completeness gap."),
        metric("M1.3", "Parse rate vs raw lines",
               "M1.2 / M1.1 * 100. The headline ingestion KPI.",
               round(parsed_pct, 2), "pct",
               ">= 95.0%", status_for(100 - parsed_pct, target_max=5.0, amber_max=10.0),
               "Both",
               "If this is below 95% the vendor will silently drop records. The remaining 3.30% on this dataset are continuation lines absorbed into multi-line cells (NOT lost)."),
        metric("M1.4", "Quarantine bucket size after P1",
               "Rows that the P1 pipeline could not repair and chose to quarantine for manual review.",
               quarantined, "records",
               "0", status_for(quarantined, target_max=0, amber_max=100),
               "Surveillance team",
               "Quarantined records are deferred surveillance gaps — every quarantined row is a row no rule will fire on."),
        metric("M1.5", "Bare-CR lines (silent-drop risk)",
               "Lines terminated by a bare \\r rather than \\r\\n. Strict-LF readers (some legacy Java/C parsers, Unix wc) silently miss these.",
               bare_cr, "records",
               "0 in the FUTURE (producer should emit CRLF only)", "AMBER", "Producer",
               "On THIS file the P1 pipeline normalises bare-CR -> CRLF, so no records are lost downstream. But the producer should fix the line-ending mix at source so every consumer sees the same record count."),
        metric("M1.6", "CRLF-terminated lines",
               "Lines terminated by \\r\\n (standard Windows + RFC-4180 expectation).",
               crlf, "records",
               "n/a (input characteristic)", "GREEN", "Producer",
               "Reported alongside M1.5 so a reader can reconcile any line-counting tool against the file."),
    ]
    out["themes"]["M1 — Ingestion completeness"] = {"metrics": m1}

    # ============================================================
    # M2 — Field-level completeness
    # ============================================================
    # Critical fields where null = surveillance gap; optional fields where null is by design
    field_policy = {
        # field            : (target_null_pct_max, amber_null_pct_max, owner, is_optional, rationale)
        "ExchangeId":         (1.0, 5.0, "Producer", False, "Required for venue-level surveillance + best-execution analysis."),
        "MessageType":        (0.0, 0.0, "Producer", False, "Cannot route a row to a surveillance rule without knowing what kind of event it is."),
        "TransactionTime":    (0.0, 0.0, "Producer", False, "Required for every time-windowed rule (cancel ratios, wash-trade timing, lifecycle ordering)."),
        "MessageDate":        (0.0, 0.0, "Producer", False, "Partition key for trading-day surveillance scope."),
        "MessageId":          (0.0, 1.0, "Producer", False, "Primary key. Nulls cannot be deduplicated, reconciled, or referenced by Cancel/Fill events."),
        "LinkMessageId":      (None, None, "Producer", True, "Optional — null is valid for parent New Orders. Only flagged when populated AND unresolvable (see M3.1)."),
        "ParentOrderId":      (None, None, "Producer", True, "Optional — null is valid for parent New Orders. Only flagged when populated AND unresolvable (see M3.2)."),
        "Instrument":         (0.0, 0.0, "Producer", False, "Required for every per-instrument rule (price outlier, position aggregation)."),
        "ISIN":               (0.0, 1.0, "Producer", False, "Required for cross-venue / cross-system instrument joins."),
        "BuyOrSell":          (0.0, 0.0, "Producer", False, "Required for wash-trade detection, position direction, side-imbalance rules."),
        "Price":              (1.0, 5.0, "Producer", False, "Required for price-outlier detection + best-execution. Some Cancel events may legitimately omit price; investigate the breakdown."),
        "TotalVolume":        (1.0, 5.0, "Producer", False, "Required for over-fill detection + position size + block-trade flagging."),
        "Account":            (0.0, 0.0, "Producer", False, "Required for account-level wash-trade and position-limit rules."),
        "CounterPartyFirm":   (0.0, 0.0, "Producer", False, "Required for inter-firm market-abuse detection + counterparty risk."),
        "Trader":             (0.0, 5.0, "Producer", False, "Reg-best-execution + MAR market-abuse attribution. Currently 64.87% of FILLS lack a Trader (see M6)."),
        "TransactionSource":  (0.0, 0.0, "Producer", False, "Required for source attribution if the surveillance system aggregates multiple feeds."),
        "Currency":           (0.0, 0.0, "Producer", False, "Required for FX normalisation if multi-currency."),
        "Flags":              (None, None, "Producer", True, "Optional — present only on exception messages. Null is the expected state on the majority of rows."),
    }

    m2 = []
    for col, (tgt_max, amb_max, owner, is_opt, rationale) in field_policy.items():
        if col not in profile["columns"]:
            continue
        n_null = profile["columns"][col]["p1"]["n_null"]
        pct_null = round(n_null / n * 100, 2)
        if is_opt:
            status = "GREEN"
            target_str = "n/a (optional field)"
        else:
            status = status_for(pct_null, target_max=tgt_max, amber_max=amb_max)
            target_str = f"<= {tgt_max}% null"
        m2.append(metric(
            f"M2.{col}", f"{col} populated rate",
            f"100 - (nulls in {col} / total rows * 100)",
            round(100 - pct_null, 2), "pct",
            target_str if is_opt else f">= {round(100 - tgt_max, 2)}% populated",
            status, owner, rationale,
        ))
    out["themes"]["M2 — Field-level completeness"] = {"metrics": m2}

    # ============================================================
    # M3 — Lineage integrity
    # ============================================================
    new_orders = df[(df["MessageType"] == "New Order")]
    n_new_orders = len(new_orders)
    a4_pct = round(p2["A4"]["count"] / max(n_new_orders, 1) * 100, 2)

    m3 = [
        metric("M3.1", "Orphaned LinkMessageId rate",
               "P2 finding A1 / count(LinkMessageId is non-null) * 100. References to MessageIds that don't exist in the file.",
               p2["A1"]["pct_of_link_set"], "pct",
               "< 1.0%", status_for(p2["A1"]["pct_of_link_set"], target_max=1.0, amber_max=5.0),
               "Producer",
               "Every orphaned reference is a child event the surveillance system cannot tie to its order. Lifecycle reconstruction breaks for those rows."),
        metric("M3.2", "Orphaned ParentOrderId rate",
               "P2 finding A2 / count(ParentOrderId is non-null) * 100. References to MessageIds that don't exist in the file.",
               p2["A2"]["pct_of_parent_set"], "pct",
               "< 1.0%", status_for(p2["A2"]["pct_of_parent_set"], target_max=1.0, amber_max=5.0),
               "Producer",
               "Breaks the textbook 'aggregate Fill <= parent volume' check (see M4.2). Producer must guarantee parent records are emitted before children."),
        metric("M3.3", "Orphan-by-omission rate (children with NO upstream ref)",
               "P2 finding A3 / count(child events) * 100. Cancel/Replace/Fill/Reject rows where BOTH LinkMessageId AND ParentOrderId are null.",
               p2["A3"].get("pct_of_child_messages", 0.0), "pct",
               "0.0%", status_for(p2["A3"].get("pct_of_child_messages", 0.0), target_max=0.0, amber_max=0.5),
               "Producer",
               "Hardest class of completeness gap — these rows cannot be tied to ANY order. A regulator audit will fail on these."),
        metric("M3.4", "Schema misuse rate (New Order with parent)",
               "P2 finding A4 / count(New Order rows) * 100. New Orders carrying LinkMessageId or ParentOrderId.",
               a4_pct, "pct",
               "0.0%", status_for(a4_pct, target_max=0.0, amber_max=2.0),
               "Producer",
               "Either MessageType is mislabelled or the parent column is mis-populated. Either way the surveillance system cannot trust the type label on these rows."),
    ]
    out["themes"]["M3 — Lineage integrity"] = {"metrics": m3}

    # ============================================================
    # M4 — Lifecycle coherence
    # ============================================================
    # B2 denominator: parents that have at least one fill
    is_fill = df["MessageType"] == "Fill"
    fills_parent_refs = df.loc[is_fill, "ParentOrderId"].dropna().unique()
    parents_with_fills = len(fills_parent_refs)
    b2_pct = round(p2["B2"]["count"] / max(parents_with_fills, 1) * 100, 2)

    m4 = [
        metric("M4.1", "Time-travel rate (child before parent)",
               "P2 finding B1 / count(children with resolvable parent) * 100. Child event timestamp < parent timestamp.",
               p2["B1"]["pct_of_resolvable_children"], "pct",
               "0.0%", status_for(p2["B1"]["pct_of_resolvable_children"], target_max=0.0, amber_max=1.0),
               "Producer",
               "Physically impossible. Indicates either timestamp corruption (clock skew at the producer) or misassigned parent reference."),
        metric("M4.2", "Over-fill rate (parents where fills > order volume)",
               "P2 finding B2 / count(parents with at least one fill) * 100.",
               b2_pct, "pct",
               "0.0%", status_for(b2_pct, target_max=0.0, amber_max=1.0),
               "Producer",
               f"Median over-fill on this dataset is {p2['B2'].get('median_over_fill_pct', 0)}% (plausible rounding); MAX is {p2['B2'].get('max_over_fill_pct', 0)}% (impossible — corruption)."),
        metric("M4.3", "Fill-after-cancel rows",
               "P2 finding B3. Fill on a parent that was already Cancelled.",
               p2["B3"]["count"], "records",
               "Investigate (real edge cases exist)", "AMBER", "Surveillance team",
               "On REAL data this is a surveillance alert (race condition or fabricated fill). On synthetic data the count is high — likely seeded. The rule is wired correctly either way."),
        metric("M4.4", "Replace-without-parent rate",
               "P2 finding B4 / count(Replace Order rows) * 100.",
               p2["B4"]["pct_of_replaces"], "pct",
               "0.0%", status_for(p2["B4"]["pct_of_replaces"], target_max=0.0, amber_max=2.0),
               "Producer",
               "Replace is meaningless without the order it's replacing — surveillance can't compare old vs new price/quantity."),
    ]
    out["themes"]["M4 — Lifecycle coherence"] = {"metrics": m4}

    # ============================================================
    # M5 — Reference data stability
    # ============================================================
    c1_pct = round(p2["C1"]["count"] / p2["C1"]["total_instruments"] * 100, 2)
    c2_pct = round(p2["C2"]["count"] / max(p2["C2"]["total_isins"], 1) * 100, 2)
    d1_pct = round(p2["D1"]["count"] / n * 100, 2)
    # PROP vs non-PROP self-cross split
    prop_self = p2["D1"]["prefix_breakdown"].get("PROP", 0)
    non_prop_self = p2["D1"]["count"] - prop_self
    non_prop_self_pct = round(non_prop_self / n * 100, 2)

    m5 = [
        metric("M5.1", "Instrument->ISIN drift rate",
               "P2 finding C1 / total distinct instruments * 100. Instruments mapped to multiple ISINs.",
               c1_pct, "pct",
               "< 1.0%", status_for(c1_pct, target_max=1.0, amber_max=10.0),
               "Producer",
               "HEADLINE DQ FINDING. Per-instrument surveillance rules (price outlier, position aggregation, market-data join) cannot fire reliably until this is resolved."),
        metric("M5.2", "ISIN->Instrument drift rate",
               "P2 finding C2 / total distinct ISINs * 100. Reverse mapping.",
               c2_pct, "pct",
               "0.0%", status_for(c2_pct, target_max=0.0, amber_max=1.0),
               "Producer",
               "Reverse mapping is stable on this dataset (0%). Reported here so the asymmetry with M5.1 is documented."),
        metric("M5.3", "Account == CounterPartyFirm collapse rate",
               "P2 finding D1 / total rows * 100. Account string identical to CPF string.",
               d1_pct, "pct",
               "< 5.0% (PROP-only is acceptable)", status_for(d1_pct, target_max=5.0, amber_max=20.0),
               "Producer",
               f"Stage-1 profile shows Account has 1,043 distinct values vs CPF's 100, so granular accounts EXIST but are absent on {d1_pct}% of rows. Looks like an upstream column-copy bug."),
        metric("M5.4", "Non-PROP self-cross rate",
               "Subset of M5.3 where Account prefix is not PROP. These are the wash-trade-relevant ones.",
               non_prop_self_pct, "pct",
               "< 0.5%", status_for(non_prop_self_pct, target_max=0.5, amber_max=5.0),
               "Producer",
               "PROP self-cross is legit; CLNT/BRKR/INST self-cross should be reviewed. Currently masked by M5.3 — fix M5.3 first."),
    ]
    out["themes"]["M5 — Reference data stability"] = {"metrics": m5}

    # ============================================================
    # M6 — Attribution coverage
    # ============================================================
    trader_by_type = p2["E1"]["trader_pct_by_msgtype"]
    fill_with_trader_pct = trader_by_type.get("Fill", {}).get("with_trader_pct", 0)
    new_with_trader_pct = trader_by_type.get("New Order", {}).get("with_trader_pct", 0)
    cancel_with_trader_pct = trader_by_type.get("Cancel Order", {}).get("with_trader_pct", 0)
    replace_with_trader_pct = trader_by_type.get("Replace Order", {}).get("with_trader_pct", 0)

    m6 = [
        metric("M6.1", "Fill attribution rate (% of fills with Trader)",
               "100 - (E1 / count(Fill) * 100). Coverage of Reg-best-execution / MAR attribution.",
               fill_with_trader_pct, "pct",
               ">= 95.0%", status_for(fill_with_trader_pct, target_min=95.0, amber_min=80.0),
               "Producer",
               "REGULATOR-EXAM CRITICAL. Currently 35.13% — the surveillance system cannot attribute 65% of fills to a human."),
        metric("M6.2", "New Order attribution rate",
               "% of New Order rows with Trader populated.",
               new_with_trader_pct, "pct",
               ">= 95.0%", status_for(new_with_trader_pct, target_min=95.0, amber_min=80.0),
               "Producer",
               "Required for spoofing/layering attribution at the parent level."),
        metric("M6.3", "Cancel Order attribution rate",
               "% of Cancel Order rows with Trader populated.",
               cancel_with_trader_pct, "pct",
               ">= 95.0%", status_for(cancel_with_trader_pct, target_min=95.0, amber_min=80.0),
               "Producer",
               "Required for cancel-ratio-per-trader (F1) to be meaningful."),
        metric("M6.4", "Replace Order attribution rate",
               "% of Replace Order rows with Trader populated.",
               replace_with_trader_pct, "pct",
               ">= 95.0%", status_for(replace_with_trader_pct, target_min=95.0, amber_min=80.0),
               "Producer",
               "Required for modify-then-spoof / quote-stuffing detection."),
    ]
    out["themes"]["M6 — Attribution coverage"] = {"metrics": m6}

    # ============================================================
    # M7 — Surveillance rule eligibility (eligible-population sizes)
    # ============================================================
    n_fills = int(is_fill.sum())
    # F2 eligible population: fills with non-null Account + Instrument + parseable timestamp
    df_ts = pd.to_datetime(df["TransactionTime"], format="%Y-%m-%dT%H:%M:%S.%fZ", errors="coerce", utc=True)
    f2_eligible = int(((df["MessageType"] == "Fill") & df["Account"].notna() & df["Instrument"].notna() & df_ts.notna()).sum())
    f2_eligible_pct = round(f2_eligible / max(n_fills, 1) * 100, 2)

    # G1 eligible population: rows with Price > 0 + Instrument
    df["_p"] = pd.to_numeric(df["Price"], errors="coerce")
    g1_eligible = int((df["_p"].notna() & (df["_p"] > 0) & df["Instrument"].notna()).sum())
    g1_eligible_pct = round(g1_eligible / n * 100, 2)

    # B2 eligible: parents with TotalVolume + at least one fill with parseable volume
    new_with_vol = df[(df["MessageType"] == "New Order") & df["TotalVolume"].notna()]["MessageId"].dropna().unique()
    fills_with_parent = df[is_fill & df["ParentOrderId"].notna() & df["TotalVolume"].notna()]
    b2_eligible_parents = len(set(fills_with_parent["ParentOrderId"]) & set(new_with_vol))
    n_new_with_vol = len(new_with_vol)
    b2_eligible_pct = round(b2_eligible_parents / max(n_new_with_vol, 1) * 100, 2)

    m7 = [
        metric("M7.1", "F2 (wash-trade) eligible-population rate",
               "Fills with non-null Account + Instrument + parseable timestamp / total Fills * 100.",
               f2_eligible_pct, "pct",
               ">= 95.0%", status_for(f2_eligible_pct, target_min=95.0, amber_min=80.0),
               "Surveillance team",
               f"Currently {f2_eligible:,} of {n_fills:,} fills are eligible. The wash-trade rule cannot fire on the rest."),
        metric("M7.2", "G1 (price outlier) eligible-population rate",
               "Rows with Price > 0 + non-null Instrument / total rows * 100.",
               g1_eligible_pct, "pct",
               ">= 95.0%", status_for(g1_eligible_pct, target_min=95.0, amber_min=80.0),
               "Surveillance team",
               "Per-instrument outlier detection has high coverage, but the per-instrument distribution is unreliable until M5.1 (Instrument->ISIN drift) is resolved."),
        metric("M7.3", "B2 (over-fill) eligible-parent rate",
               "Parents with TotalVolume + at least one fill with parseable volume / parents with TotalVolume * 100.",
               b2_eligible_pct, "pct",
               ">= 80.0%", status_for(b2_eligible_pct, target_min=80.0, amber_min=50.0),
               "Surveillance team",
               "Parents without fills are ineligible by definition — this metric only counts the rule-relevant denominator."),
        metric("M7.4", "F1 (cancel ratio) eligible-trader population",
               "Distinct Traders with at least one New Order AND at least one Cancel Order.",
               len(p2["F1"]["top_10_traders_by_cancel_ratio"]), "traders (top-10 reported)",
               "All active traders eligible", "AMBER", "Surveillance team",
               "F1 is only useful with a tuned threshold AND high Trader coverage (see M6). On this dataset top ratios are 0.68-0.72 — no spoofing pattern seeded."),
    ]
    out["themes"]["M7 — Surveillance rule eligibility"] = {"metrics": m7}

    # ============================================================
    # Top-line scorecard summary
    # ============================================================
    all_metrics = [m for theme in out["themes"].values() for m in theme["metrics"]]
    by_status: dict[str, int] = {"GREEN": 0, "AMBER": 0, "RED": 0}
    for m in all_metrics:
        by_status[m["status"]] = by_status.get(m["status"], 0) + 1
    out["summary"] = {
        "total_metrics": len(all_metrics),
        "by_status": by_status,
        "headline": {
            "ingestion_parse_rate_pct": parsed_pct,
            "fill_attribution_rate_pct": fill_with_trader_pct,
            "instrument_isin_drift_pct": c1_pct,
            "account_cpf_collapse_pct": d1_pct,
            "lineage_orphan_parent_pct": p2["A2"]["pct_of_parent_set"],
        },
    }

    OUT.write_text(json.dumps(out, indent=2, default=str), encoding="utf-8")
    print(f"\nWrote: {OUT.name}")
    print(f"\nSummary: {len(all_metrics)} metrics across {len(out['themes'])} themes")
    print(f"  GREEN: {by_status['GREEN']}")
    print(f"  AMBER: {by_status['AMBER']}")
    print(f"  RED:   {by_status['RED']}")
    print("\nHeadline KPIs:")
    for k, v in out["summary"]["headline"].items():
        print(f"  {k}: {v}")


if __name__ == "__main__":
    main()
