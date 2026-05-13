"""
Test suite for the P2 surveillance-domain findings.

Every finding count cited in stage_p2_accuracy_efficacy.docx is re-derived
from cleaned_p1_final.tsv and asserted against the value frozen in
09_p2_findings.json. If a count drifts -- because the data changed, the
detection logic changed, or the report cited the wrong number -- a test
breaks.

This is the P2 analogue of 12_test_cleaned_p1.py:
  - 12 verifies the P1 PIPELINE'S INVARIANTS hold on the output file
  - 13 verifies the P2 REPORT'S COUNTS reproduce from the output file

Together they make every number in the deliverable interrogable.

Exit code: 0 if every finding reproduces, non-zero otherwise.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
SRC = HERE / "cleaned_p1_final.tsv"
P2_FINDINGS = HERE / "09_p2_findings.json"

CHILD_TYPES = {"Cancel Order", "Replace Order", "Fill", "Reject"}
WASH_WINDOW_SEC = 5

results: list[tuple[str, str, str, str]] = []


def record(test_id: str, name: str, ok: bool, detail: str) -> None:
    results.append((test_id, name, "PASS" if ok else "FAIL", detail))
    marker = "[OK]  " if ok else "[FAIL]"
    print(f"  {marker} {test_id:6s} {name}")
    if detail:
        for line in detail.splitlines():
            print(f"           {line}")


def section(title: str) -> None:
    print(f"\n{title}")
    print("-" * len(title))


def assert_count(test_id: str, name: str, actual: int, expected: int) -> None:
    record(test_id, name, actual == expected, f"actual={actual:,}  expected={expected:,}")


def assert_pct(test_id: str, name: str, actual: float, expected: float, tol: float = 0.05) -> None:
    record(test_id, name, abs(actual - expected) <= tol, f"actual={actual:.2f}%  expected={expected:.2f}%  tol=+/-{tol}%")


def main() -> int:
    print(f"Loading {SRC.name} ...")
    df = pd.read_csv(SRC, sep="\t", dtype=str, keep_default_na=True, na_values=[""])
    findings = json.loads(P2_FINDINGS.read_text(encoding="utf-8"))
    n = len(df)
    print(f"Loaded: {n:,} rows")

    # Parse timestamps for time-based checks (must come BEFORE the parent lookup so it carries _ts)
    df["_ts"] = pd.to_datetime(df["TransactionTime"], format="%Y-%m-%dT%H:%M:%S.%fZ", errors="coerce", utc=True)

    # Reference set: every non-null MessageId that exists in the file (used for A1/A2 orphan checks)
    mid_set = set(df.loc[df["MessageId"].notna(), "MessageId"])

    # Parent lookup: New Orders ONLY (matches 09_p2_accuracy_efficacy.py).
    # The B-theme checks use this restricted lookup because only New Orders are valid parents.
    new_orders_all = df[(df["MessageType"] == "New Order") & df["MessageId"].notna()]
    parent_lookup = new_orders_all.drop_duplicates(subset=["MessageId"], keep="first").set_index("MessageId")
    parent_ts_map = parent_lookup["_ts"]
    parent_vol_map = pd.to_numeric(parent_lookup["TotalVolume"], errors="coerce")

    # Resolve each row's parent reference: prefer ParentOrderId, fall back to LinkMessageId
    df["_parent_ref"] = df["ParentOrderId"].where(df["ParentOrderId"].notna(), df["LinkMessageId"])
    df["_parent_ts"] = df["_parent_ref"].map(parent_ts_map)
    df["_parent_vol"] = df["_parent_ref"].map(parent_vol_map)

    # ============================================================
    # A. Lineage
    # ============================================================
    section("A. Lineage")

    # A1 -- orphaned LinkMessageId
    has_link = df["LinkMessageId"].notna()
    a1_orphan = has_link & ~df["LinkMessageId"].isin(mid_set)
    assert_count("A1", "Orphaned LinkMessageId rows", int(a1_orphan.sum()), findings["A1"]["count"])
    a1_pct = a1_orphan.sum() / has_link.sum() * 100
    assert_pct("A1.pct", "A1 % of link-set", a1_pct, findings["A1"]["pct_of_link_set"])

    # A2 -- orphaned ParentOrderId
    has_parent = df["ParentOrderId"].notna()
    a2_orphan = has_parent & ~df["ParentOrderId"].isin(mid_set)
    assert_count("A2", "Orphaned ParentOrderId rows", int(a2_orphan.sum()), findings["A2"]["count"])
    a2_pct = a2_orphan.sum() / has_parent.sum() * 100
    assert_pct("A2.pct", "A2 % of parent-set", a2_pct, findings["A2"]["pct_of_parent_set"])

    # A3 -- child message with NO upstream ref
    is_child = df["MessageType"].isin(CHILD_TYPES)
    a3 = is_child & df["LinkMessageId"].isna() & df["ParentOrderId"].isna()
    assert_count("A3", "Child message with NO upstream ref", int(a3.sum()), findings["A3"]["count"])

    # A4 -- New Order carrying upstream ref
    is_new = df["MessageType"] == "New Order"
    a4 = is_new & (df["LinkMessageId"].notna() | df["ParentOrderId"].notna())
    assert_count("A4", "New Order with parent/link reference", int(a4.sum()), findings["A4"]["count"])

    # ============================================================
    # B. Order lifecycle
    # ============================================================
    section("B. Order lifecycle")

    # B1 -- child timestamp BEFORE referenced parent's timestamp
    have_parent_ts = df["_parent_ts"].notna() & df["_ts"].notna()
    b1 = have_parent_ts & (df["_ts"] < df["_parent_ts"])
    assert_count("B1", "Child event BEFORE parent", int(b1.sum()), findings["B1"]["count"])

    # B2 -- aggregate Fill volume per parent > parent TotalVolume
    is_fill = df["MessageType"] == "Fill"
    fills_with_parent = df[is_fill & df["_parent_ref"].notna()].copy()
    fills_with_parent["TotalVolume_num"] = pd.to_numeric(fills_with_parent["TotalVolume"], errors="coerce")
    fill_sum = fills_with_parent.groupby("_parent_ref")["TotalVolume_num"].sum()
    parent_compare = pd.concat(
        [fill_sum.rename("filled"), parent_vol_map.rename("order_vol")],
        axis=1, join="inner"
    ).dropna()
    over_filled = parent_compare[parent_compare["filled"] > parent_compare["order_vol"]]
    assert_count("B2", "Parents where aggregate fill > order volume", len(over_filled), findings["B2"]["count"])

    # B3 -- Fill recorded AFTER parent's earliest Cancel
    cancel_min = (
        df[df["MessageType"] == "Cancel Order"]
          .dropna(subset=["_parent_ref", "_ts"])
          .groupby("_parent_ref")["_ts"].min()
    )
    fills_with_parent["_cancel_ts"] = fills_with_parent["_parent_ref"].map(cancel_min)
    b3 = (fills_with_parent["_ts"] > fills_with_parent["_cancel_ts"]).fillna(False)
    assert_count("B3", "Fill recorded AFTER parent cancelled", int(b3.sum()), findings["B3"]["count"])

    # B4 -- Replace Order whose parent reference doesn't resolve to a New Order
    is_replace = df["MessageType"] == "Replace Order"
    b4 = is_replace & df["_parent_ts"].isna()
    assert_count("B4", "Replace Order without resolvable parent", int(b4.sum()), findings["B4"]["count"])

    # ============================================================
    # C. Instrument / ISIN
    # ============================================================
    section("C. Instrument / ISIN")

    inst_isin = df.dropna(subset=["Instrument", "ISIN"])
    per_inst_nunique = inst_isin.groupby("Instrument")["ISIN"].nunique()
    c1 = (per_inst_nunique > 1).sum()
    assert_count("C1", "Instruments mapped to multiple ISINs", int(c1), findings["C1"]["count"])
    record("C1.uni", "C1 covers expected instrument universe size",
           per_inst_nunique.shape[0] == findings["C1"]["total_instruments"],
           f"actual={per_inst_nunique.shape[0]}  expected={findings['C1']['total_instruments']}")

    per_isin_nunique = inst_isin.groupby("ISIN")["Instrument"].nunique()
    c2 = (per_isin_nunique > 1).sum()
    assert_count("C2", "ISINs mapped to multiple Instruments", int(c2), findings["C2"]["count"])

    # ============================================================
    # D. Account / Firm
    # ============================================================
    section("D. Account / Firm")

    both_present = df["Account"].notna() & df["CounterPartyFirm"].notna()
    d1 = both_present & (df["Account"] == df["CounterPartyFirm"])
    assert_count("D1", "Account == CounterPartyFirm (self-trade signature)", int(d1.sum()), findings["D1"]["count"])

    # D1b -- non-PROP self-cross
    d1b = d1 & ~df["Account"].fillna("").str.startswith("PROP")
    assert_count("D1b", "Non-PROP self-cross", int(d1b.sum()), findings["D1b"]["count"])

    # ============================================================
    # E. Trader
    # ============================================================
    section("E. Trader")

    e1 = is_fill & df["Trader"].isna()
    assert_count("E1", "Fill without attributed Trader", int(e1.sum()), findings["E1"]["count"])
    e1_pct = e1.sum() / is_fill.sum() * 100
    assert_pct("E1.pct", "E1 % of all fills", e1_pct, findings["E1"]["pct_of_fills"])

    # ============================================================
    # F. Surveillance signatures
    # ============================================================
    section("F. Surveillance signatures")

    # F1 -- cancel-to-new-order ratio per Trader (top 10)
    by_trader = df.dropna(subset=["Trader"]).groupby(["Trader", "MessageType"]).size().unstack(fill_value=0)
    by_trader["ratio"] = by_trader.get("Cancel Order", 0) / by_trader.get("New Order", 0).replace(0, pd.NA)
    top10_actual = by_trader.dropna(subset=["ratio"]).nlargest(10, "ratio")
    top10_expected = findings["F1"]["top_10_traders_by_cancel_ratio"]
    f1_match = True
    f1_detail_lines = []
    for trader, exp in top10_expected.items():
        if trader not in top10_actual.index:
            f1_match = False
            f1_detail_lines.append(f"missing {trader} from actual top-10")
            continue
        actual_ratio = round(float(top10_actual.loc[trader, "ratio"]), 2)
        if actual_ratio != exp["ratio"]:
            f1_match = False
            f1_detail_lines.append(f"{trader}: actual={actual_ratio} expected={exp['ratio']}")
    record("F1", "F1 top-10 trader cancel ratios reproduce", f1_match,
           "\n".join(f1_detail_lines) if f1_detail_lines else f"{len(top10_expected)} traders match")

    # F2 -- wash-trade signature (same Account, opposite side, <= 5s)
    fills_for_f2 = df[is_fill & df["Account"].notna() & df["Instrument"].notna() & df["_ts"].notna()].copy()
    fills_for_f2 = fills_for_f2.sort_values(["Account", "Instrument", "_ts"])
    fills_for_f2["_prev_side"] = fills_for_f2.groupby(["Account", "Instrument"])["BuyOrSell"].shift(1)
    fills_for_f2["_prev_ts"] = fills_for_f2.groupby(["Account", "Instrument"])["_ts"].shift(1)
    fills_for_f2["_dt"] = (fills_for_f2["_ts"] - fills_for_f2["_prev_ts"]).dt.total_seconds()
    f2_mask = (
        fills_for_f2["_prev_side"].notna()
        & (fills_for_f2["_prev_side"] != fills_for_f2["BuyOrSell"])
        & (fills_for_f2["_dt"] <= WASH_WINDOW_SEC)
    )
    assert_count("F2", "Wash-trade signature rows", int(f2_mask.sum()), findings["F2"]["count"])

    # ============================================================
    # G. Statistical
    # ============================================================
    section("G. Statistical outliers")

    # G1 runs on ALL rows with price > 0 + non-null Instrument (not just fills) -- matches 09_p2_accuracy_efficacy.py
    df["Price_num"] = pd.to_numeric(df["Price"], errors="coerce")
    valid_px = df[df["Price_num"].notna() & (df["Price_num"] > 0) & df["Instrument"].notna()].copy()
    q1 = valid_px.groupby("Instrument")["Price_num"].quantile(0.25)
    q3 = valid_px.groupby("Instrument")["Price_num"].quantile(0.75)
    iqr = q3 - q1
    valid_px["_lo"] = valid_px["Instrument"].map(q1 - 3 * iqr)
    valid_px["_hi"] = valid_px["Instrument"].map(q3 + 3 * iqr)
    g1_mask = (valid_px["Price_num"] < valid_px["_lo"]) | (valid_px["Price_num"] > valid_px["_hi"])
    assert_count("G1", "Price outliers (per-instrument 3xIQR)", int(g1_mask.sum()), findings["G1"]["count"])

    df["Vol_num"] = pd.to_numeric(df["TotalVolume"], errors="coerce")
    valid_vol = df[df["Vol_num"].notna() & (df["Vol_num"] > 0) & df["Instrument"].notna()].copy()
    vq1 = valid_vol.groupby("Instrument")["Vol_num"].quantile(0.25)
    vq3 = valid_vol.groupby("Instrument")["Vol_num"].quantile(0.75)
    viqr = vq3 - vq1
    valid_vol["_vhi"] = valid_vol["Instrument"].map(vq3 + 3 * viqr)
    g2_mask = valid_vol["Vol_num"] > valid_vol["_vhi"]
    assert_count("G2", "Volume outliers (per-instrument upper 3xIQR)", int(g2_mask.sum()), findings["G2"]["count"])

    # ============================================================
    # SUMMARY
    # ============================================================
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    n_pass = sum(1 for _, _, s, _ in results if s == "PASS")
    n_fail = sum(1 for _, _, s, _ in results if s == "FAIL")
    print(f"  Total tests:  {len(results)}")
    print(f"  Passed:       {n_pass}")
    print(f"  Failed:       {n_fail}")
    if n_fail:
        print("\n  Failures:")
        for tid, name, status, detail in results:
            if status == "FAIL":
                print(f"    - {tid}  {name}")
                print(f"        {detail}")
    return 0 if n_fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
