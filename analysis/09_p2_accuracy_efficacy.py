"""
Stage P2 — Accuracy & efficacy checks (surveillance-domain lens).

The brief explicitly frames this priority as open-ended ("based on your imagination").
The lens here is surveillance: "can the data, even if mechanically clean, support
the surveillance rules a regulator expects us to run?" Each check is named, has a
stated surveillance rationale, and produces both an aggregate count and (where
useful) a small set of row-level examples.

Inputs:
  - cleaned_p1_final.tsv (output of Stage P1)
  - 06_row_flags.tsv (per-row P1 flags)

Outputs:
  - 09_p2_findings.json — every check's findings (rationale + count + examples)
  - 09_p2_examples.json — row-level evidence per check
"""

from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
SRC = HERE / "cleaned_p1_final.tsv"
FLAGS_SRC = HERE / "06_row_flags.tsv"
FINDINGS = HERE / "09_p2_findings.json"
EXAMPLES = HERE / "09_p2_examples.json"


def load() -> tuple[pd.DataFrame, pd.DataFrame]:
    print("Loading cleaned P1 data ...")
    df = pd.read_csv(SRC, sep="\t", dtype=str, keep_default_na=False, na_filter=False)
    # Coerce numerics
    df["Price_num"] = pd.to_numeric(df["Price"], errors="coerce")
    df["TotalVolume_num"] = pd.to_numeric(df["TotalVolume"], errors="coerce")
    df["TransactionTime_ts"] = pd.to_datetime(df["TransactionTime"], format="%Y-%m-%dT%H:%M:%S.%fZ", errors="coerce", utc=True)
    flags = pd.read_csv(FLAGS_SRC, sep="\t")
    print(f"  Rows: {len(df):,}")
    return df, flags


def add_finding(out: dict, key: str, *, theme: str, name: str, rationale: str,
                detection: str, count: int, extras: dict | None = None) -> None:
    out[key] = {
        "theme": theme,
        "name": name,
        "rationale": rationale,
        "detection": detection,
        "count": int(count),
        **(extras or {}),
    }
    print(f"  [{key}] {name}: {count:,}")


def sample_rows(df: pd.DataFrame, mask, k: int = 5, cols: list[str] | None = None):
    cols = cols or ["MessageId", "MessageType", "TransactionTime", "Instrument", "BuyOrSell",
                    "Price", "TotalVolume", "Account", "CounterPartyFirm", "Trader",
                    "LinkMessageId", "ParentOrderId"]
    idx = df.index[mask][:k]
    return df.loc[idx, cols].astype(str).to_dict(orient="records")


def main() -> None:
    df, flags = load()
    findings: dict = {}
    examples: dict = {}

    # Useful subsets
    by_type = df.groupby("MessageType", dropna=False).size().to_dict()
    print("\nMessageType distribution:")
    for k, v in sorted(by_type.items(), key=lambda kv: -kv[1]):
        print(f"  {k!r:25} {v:,}")

    # Build sets of valid PKs to test references against
    valid_pks = set(df.loc[df["MessageId"] != "", "MessageId"])
    print(f"\nUnique non-null MessageIds: {len(valid_pks):,}")

    # ============================================================
    # THEME A — Referential integrity (lineage)
    # ============================================================
    print("\n=== THEME A — Referential integrity ===")

    # A1: LinkMessageId set but doesn't exist as a MessageId anywhere
    link_set = df["LinkMessageId"] != ""
    link_unknown = link_set & ~df["LinkMessageId"].isin(valid_pks)
    add_finding(findings, "A1", theme="A. Lineage",
                name="Orphaned LinkMessageId (references a MessageId that doesn't exist in the file)",
                rationale="Cancel/Replace/Fill events reference an upstream message via LinkMessageId. If that reference is unresolvable, the surveillance system cannot reconstruct the order lifecycle and any rule that depends on the chain (e.g., 'cancels before fills') silently fails.",
                detection="LinkMessageId is non-empty AND not present in the set of MessageIds in this file.",
                count=int(link_unknown.sum()),
                extras={"pct_of_link_set": round(100 * link_unknown.sum() / max(link_set.sum(), 1), 2)})
    examples["A1"] = sample_rows(df, link_unknown, k=5)

    # A2: ParentOrderId set but doesn't exist as a MessageId
    parent_set = df["ParentOrderId"] != ""
    parent_unknown = parent_set & ~df["ParentOrderId"].isin(valid_pks)
    add_finding(findings, "A2", theme="A. Lineage",
                name="Orphaned ParentOrderId (references a MessageId that doesn't exist in the file)",
                rationale="ParentOrderId points to the root parent New Order. A missing parent breaks aggregate-volume checks ('Fill total <= parent volume'), which is the textbook surveillance check for execution integrity.",
                detection="ParentOrderId is non-empty AND not present in MessageIds.",
                count=int(parent_unknown.sum()),
                extras={"pct_of_parent_set": round(100 * parent_unknown.sum() / max(parent_set.sum(), 1), 2)})
    examples["A2"] = sample_rows(df, parent_unknown, k=5)

    # A3: Child messages (Cancel/Replace/Fill/Reject) with NO upstream reference at all
    child_types = {"Cancel Order", "Replace Order", "Fill", "Reject"}
    is_child = df["MessageType"].isin(child_types)
    no_link_no_parent = is_child & (df["LinkMessageId"] == "") & (df["ParentOrderId"] == "")
    add_finding(findings, "A3", theme="A. Lineage",
                name="Child message (Cancel/Replace/Fill/Reject) with NO upstream reference",
                rationale="A cancel without a referenced order, or a fill without a parent, is unattributable — surveillance cannot tell which order it belongs to. In a regulator audit this is the kind of completeness gap that fails an exam.",
                detection="MessageType ∈ {Cancel Order, Replace Order, Fill, Reject} AND both LinkMessageId and ParentOrderId are null.",
                count=int(no_link_no_parent.sum()),
                extras={"pct_of_child_messages": round(100 * no_link_no_parent.sum() / max(is_child.sum(), 1), 2)})
    examples["A3"] = sample_rows(df, no_link_no_parent, k=5)

    # A4: New Order with a parent (suspicious — parents shouldn't have parents)
    is_new = df["MessageType"] == "New Order"
    new_with_parent = is_new & ((df["LinkMessageId"] != "") | (df["ParentOrderId"] != ""))
    add_finding(findings, "A4", theme="A. Lineage",
                name="New Order carrying a LinkMessageId or ParentOrderId",
                rationale="A 'New Order' by definition has no upstream parent in this schema. Carrying a parent reference is either schema misuse or column-shift residue.",
                detection="MessageType == 'New Order' AND (LinkMessageId or ParentOrderId is non-empty).",
                count=int(new_with_parent.sum()))
    examples["A4"] = sample_rows(df, new_with_parent, k=5)

    # ============================================================
    # THEME B — Order lifecycle coherence
    # ============================================================
    print("\n=== THEME B — Order lifecycle coherence ===")

    # Build parent timestamp lookup (using New Orders as parents).
    # MessageId has 17K known duplicates (P1 finding N5c) — keep the FIRST occurrence for the lookup,
    # and record the dupe count so the doc can call it out.
    new_orders_all = df[is_new & (df["MessageId"] != "")]
    n_dup_new_pks = int(new_orders_all["MessageId"].duplicated().sum())
    new_orders = new_orders_all.drop_duplicates(subset=["MessageId"], keep="first").set_index("MessageId")
    parent_ts = new_orders["TransactionTime_ts"]
    parent_vol = new_orders["TotalVolume_num"]
    parent_instr = new_orders["Instrument"]
    print(f"  (Parent lookup built from {len(new_orders):,} unique New Order PKs; {n_dup_new_pks:,} duplicate New Order PKs collapsed)")

    # Resolve parent timestamps for child rows (prefer ParentOrderId, fall back to LinkMessageId)
    parent_ref = df["ParentOrderId"].where(df["ParentOrderId"] != "", df["LinkMessageId"])
    df["_parent_ref"] = parent_ref
    df["_parent_ts"] = df["_parent_ref"].map(parent_ts)
    df["_parent_vol"] = df["_parent_ref"].map(parent_vol)
    df["_parent_instr"] = df["_parent_ref"].map(parent_instr)

    # B1: Child event with TransactionTime BEFORE the referenced parent's TransactionTime
    have_parent_ts = df["_parent_ts"].notna() & df["TransactionTime_ts"].notna()
    before_parent = have_parent_ts & (df["TransactionTime_ts"] < df["_parent_ts"])
    add_finding(findings, "B1", theme="B. Order lifecycle",
                name="Child event occurs BEFORE its referenced parent New Order",
                rationale="Time-travel impossible — a cancel/fill cannot precede the order it acts on. Such rows indicate either timestamp corruption or a misassigned parent reference. Either way, the surveillance lifecycle reconstruction is broken for those rows.",
                detection="Child has a resolvable parent timestamp AND child.TransactionTime < parent.TransactionTime.",
                count=int(before_parent.sum()),
                extras={"pct_of_resolvable_children": round(100 * before_parent.sum() / max(have_parent_ts.sum(), 1), 2)})
    examples["B1"] = sample_rows(df, before_parent, k=5, cols=["MessageId", "MessageType", "TransactionTime", "_parent_ref", "_parent_ts", "Instrument"])

    # B2: Fill total per parent exceeds parent's TotalVolume
    is_fill = df["MessageType"] == "Fill"
    fills_with_parent = df[is_fill & df["_parent_ref"].notna() & (df["_parent_ref"] != "")].copy()
    fill_sum = fills_with_parent.groupby("_parent_ref")["TotalVolume_num"].sum()
    parent_compare = pd.concat([fill_sum.rename("filled"), parent_vol.rename("order_vol")], axis=1, join="inner").dropna()
    over_filled = parent_compare[parent_compare["filled"] > parent_compare["order_vol"]]
    add_finding(findings, "B2", theme="B. Order lifecycle",
                name="Aggregate Fill volume per parent exceeds parent's TotalVolume",
                rationale="The textbook execution-integrity check. Filling more than was ordered is impossible on a single trading line — every regulator exam will look at this. Surfacing this on the synthetic data demonstrates the check is wired.",
                detection="Sum of Fill.TotalVolume grouped by parent MessageId > parent.TotalVolume.",
                count=len(over_filled),
                extras={
                    "max_over_fill_pct": round(((over_filled["filled"] / over_filled["order_vol"]) - 1).max() * 100, 2) if len(over_filled) else 0.0,
                    "median_over_fill_pct": round(((over_filled["filled"] / over_filled["order_vol"]) - 1).median() * 100, 2) if len(over_filled) else 0.0,
                })
    if len(over_filled):
        examples["B2"] = [
            {"parent_MessageId": idx, "order_volume": float(row["order_vol"]), "filled_volume": float(row["filled"]),
             "over_by": float(row["filled"] - row["order_vol"])}
            for idx, row in over_filled.head(5).iterrows()
        ]
    else:
        examples["B2"] = []

    # B3: Fill on a previously-Cancelled parent
    cancels = df[df["MessageType"] == "Cancel Order"]
    earliest_cancel_ts = cancels.groupby("_parent_ref")["TransactionTime_ts"].min()
    fills_with_parent["_cancel_ts"] = fills_with_parent["_parent_ref"].map(earliest_cancel_ts)
    fill_after_cancel = fills_with_parent["TransactionTime_ts"] > fills_with_parent["_cancel_ts"]
    n_fac = int(fill_after_cancel.fillna(False).sum())
    add_finding(findings, "B3", theme="B. Order lifecycle",
                name="Fill recorded AFTER its parent order was Cancelled",
                rationale="A fill on a cancelled order is a surveillance red flag — it may reflect a genuine race condition (legit edge case to investigate) or fabricated execution data. Either way the surveillance system needs to know.",
                detection="Fill.TransactionTime > min(Cancel.TransactionTime for same parent).",
                count=n_fac)
    examples["B3"] = (fills_with_parent.loc[fill_after_cancel.fillna(False)].head(5)
                      [["MessageId", "MessageType", "TransactionTime", "_parent_ref", "_cancel_ts"]]
                      .astype(str).to_dict(orient="records"))

    # B4: Replace Order with no resolvable parent
    is_repl = df["MessageType"] == "Replace Order"
    repl_no_parent = is_repl & df["_parent_ts"].isna()
    add_finding(findings, "B4", theme="B. Order lifecycle",
                name="Replace Order without a resolvable parent",
                rationale="Replace is meaningless without the order it's replacing — surveillance can't compare old vs new price/quantity.",
                detection="MessageType == 'Replace Order' AND parent reference does not resolve to a parent New Order.",
                count=int(repl_no_parent.sum()),
                extras={"pct_of_replaces": round(100 * repl_no_parent.sum() / max(is_repl.sum(), 1), 2)})
    examples["B4"] = sample_rows(df, repl_no_parent, k=5)

    # ============================================================
    # THEME C — Instrument / ISIN coherence
    # ============================================================
    print("\n=== THEME C — Instrument/ISIN coherence ===")
    pairs = df[(df["Instrument"] != "") & (df["ISIN"] != "")][["Instrument", "ISIN"]]
    inst_to_isins = pairs.groupby("Instrument")["ISIN"].nunique()
    isin_to_insts = pairs.groupby("ISIN")["Instrument"].nunique()
    multi_isin_per_inst = (inst_to_isins > 1).sum()
    multi_inst_per_isin = (isin_to_insts > 1).sum()

    add_finding(findings, "C1", theme="C. Instrument/ISIN",
                name="Instrument mapped to multiple ISINs",
                rationale="A symbol-to-ISIN mapping is supposed to be stable within a trading session. Many-to-one or one-to-many breaks every per-instrument rule (price outlier, position aggregation, market-data join).",
                detection="Per Instrument: nunique(ISIN) > 1.",
                count=int(multi_isin_per_inst),
                extras={
                    "total_instruments": int(len(inst_to_isins)),
                    "top_examples": inst_to_isins.sort_values(ascending=False).head(5).to_dict(),
                })
    examples["C1"] = [{"Instrument": i, "n_distinct_ISIN": int(n)}
                      for i, n in inst_to_isins.sort_values(ascending=False).head(10).items()]

    add_finding(findings, "C2", theme="C. Instrument/ISIN",
                name="ISIN mapped to multiple Instruments",
                rationale="Same reverse problem — a single ISIN should resolve to a single ticker on a given venue/day.",
                detection="Per ISIN: nunique(Instrument) > 1.",
                count=int(multi_inst_per_isin),
                extras={
                    "total_isins": int(len(isin_to_insts)),
                    "top_examples": isin_to_insts.sort_values(ascending=False).head(5).to_dict(),
                })
    examples["C2"] = [{"ISIN": i, "n_distinct_Instrument": int(n)}
                      for i, n in isin_to_insts.sort_values(ascending=False).head(10).items()]

    # ============================================================
    # THEME D — Account / Firm relationships
    # ============================================================
    print("\n=== THEME D — Account / Firm ===")
    has_both = (df["Account"] != "") & (df["CounterPartyFirm"] != "")
    self_cross = has_both & (df["Account"] == df["CounterPartyFirm"])
    # Account prefixes signal account type
    df["_acct_prefix"] = df["Account"].str.extract(r"^([A-Z]{3,5})\d", expand=False)
    prefix_counts = df.loc[self_cross, "_acct_prefix"].value_counts().head(10).to_dict()
    add_finding(findings, "D1", theme="D. Account/Firm",
                name="Account ≡ CounterPartyFirm (self-trade signature)",
                rationale="Same Account on both sides is normal for prop trading desks (PROP prefix) but is a wash-trade signature on agency flow (CLNT, BRKR, INST prefixes). Cross-tab is the surveillance read.",
                detection="Account string identical to CounterPartyFirm string AND both non-null.",
                count=int(self_cross.sum()),
                extras={"prefix_breakdown": {str(k): int(v) for k, v in prefix_counts.items()}})
    examples["D1"] = sample_rows(df, self_cross, k=5)

    # Account prefixes that are NOT 'PROP' but still self-crossing => surveillance attention
    suspicious_self = self_cross & (~df["_acct_prefix"].fillna("").str.startswith("PROP"))
    add_finding(findings, "D1b", theme="D. Account/Firm",
                name="Non-PROP Account self-crossing (potential wash-trade signature)",
                rationale="A PROP-prefixed account self-crossing is legitimate prop activity. A CLNT/BRKR/INST account self-crossing should be reviewed — wash-trade rules trigger here.",
                detection="Self-cross rows where Account prefix is not 'PROP'.",
                count=int(suspicious_self.sum()))
    examples["D1b"] = sample_rows(df, suspicious_self, k=5)

    # ============================================================
    # THEME E — Trader attribution
    # ============================================================
    print("\n=== THEME E — Trader attribution ===")
    df["_has_trader"] = df["Trader"] != ""
    trader_by_type = df.groupby("MessageType")["_has_trader"].agg(["sum", "count"])
    trader_by_type["pct_with_trader"] = (100 * trader_by_type["sum"] / trader_by_type["count"]).round(2)

    # Fills should have a trader (someone has to be responsible for the execution)
    is_fill_mask = df["MessageType"] == "Fill"
    fill_no_trader = is_fill_mask & (df["Trader"] == "")
    add_finding(findings, "E1", theme="E. Trader",
                name="Fill without attributed Trader",
                rationale="Regulators require human attribution on executions. A Fill with no Trader is unattributable to a person and breaks Reg-best-execution / market-abuse-attribution requirements.",
                detection="MessageType == 'Fill' AND Trader is null.",
                count=int(fill_no_trader.sum()),
                extras={
                    "pct_of_fills": round(100 * fill_no_trader.sum() / max(is_fill_mask.sum(), 1), 2),
                    "trader_pct_by_msgtype": {str(idx): {"with_trader_pct": float(row["pct_with_trader"]),
                                                          "rows": int(row["count"])}
                                              for idx, row in trader_by_type.iterrows()},
                })
    examples["E1"] = sample_rows(df, fill_no_trader, k=5)

    # ============================================================
    # THEME F — Surveillance signatures
    # ============================================================
    print("\n=== THEME F — Surveillance signatures ===")

    # F1: Cancel-to-New-Order ratio per Trader (high = spoofing signature)
    type_x_trader = df[df["Trader"] != ""].groupby(["Trader", "MessageType"]).size().unstack(fill_value=0)
    has_cn = type_x_trader.reindex(columns=["New Order", "Cancel Order"], fill_value=0)
    has_cn["cancel_ratio"] = has_cn["Cancel Order"] / has_cn["New Order"].replace(0, np.nan)
    high_cancel = has_cn.sort_values("cancel_ratio", ascending=False).head(10)
    add_finding(findings, "F1", theme="F. Surveillance signatures",
                name="High cancel-to-new-order ratio per Trader",
                rationale="A persistent cancel ratio > 1 (more cancels than new orders) is the classic spoofing/layering signature — the Trader is posting orders to influence price with no intention of executing, then pulling them. Rule-of-thumb threshold: >5x for sustained activity is concerning; we surface the top-N for review.",
                detection="Group by Trader: count(Cancel Order) / count(New Order). Sort desc.",
                count=int((has_cn["cancel_ratio"] > 5).sum()),
                extras={"top_10_traders_by_cancel_ratio": {
                    str(idx): {"new_orders": int(r["New Order"]), "cancels": int(r["Cancel Order"]),
                               "ratio": (round(float(r["cancel_ratio"]), 2) if pd.notna(r["cancel_ratio"]) else None)}
                    for idx, r in high_cancel.iterrows()
                }})
    examples["F1"] = []  # report-only

    # F2: Wash-trade signature — same Account buys AND sells same Instrument within X seconds
    if df["TransactionTime_ts"].notna().any():
        side_df = df[(df["MessageType"] == "Fill") & (df["Account"] != "") & (df["Instrument"] != "") & df["TransactionTime_ts"].notna()].copy()
        side_df = side_df.sort_values(["Account", "Instrument", "TransactionTime_ts"])
        # For each Account+Instrument, look at consecutive rows where side flips and Δt is small
        side_df["_prev_side"] = side_df.groupby(["Account", "Instrument"])["BuyOrSell"].shift(1)
        side_df["_prev_ts"] = side_df.groupby(["Account", "Instrument"])["TransactionTime_ts"].shift(1)
        side_df["_dt_sec"] = (side_df["TransactionTime_ts"] - side_df["_prev_ts"]).dt.total_seconds()
        wash = side_df[(side_df["_prev_side"].notna()) & (side_df["_prev_side"] != side_df["BuyOrSell"]) & (side_df["_dt_sec"] <= 5)]
        add_finding(findings, "F2", theme="F. Surveillance signatures",
                    name="Wash-trade signature — same Account, opposite sides on same Instrument within 5s",
                    rationale="A Buy and a Sell on the same Account+Instrument within seconds is the textbook wash-trade pattern. This rule is mandatory under Reg-NMS / MiFID-II market abuse regimes.",
                    detection="Group fills by (Account, Instrument), sort by time. Flag consecutive rows where side flips AND Δtime ≤ 5s.",
                    count=len(wash),
                    extras={"pct_of_fills_with_account_and_instrument": round(100 * len(wash) / max(len(side_df), 1), 2)})
        examples["F2"] = (wash.head(5)[["Account", "Instrument", "MessageType", "BuyOrSell", "TransactionTime", "Price", "TotalVolume"]]
                          .astype(str).to_dict(orient="records"))
    else:
        add_finding(findings, "F2", theme="F. Surveillance signatures",
                    name="Wash-trade signature", rationale="Skipped — no parseable timestamps.",
                    detection="-", count=0)
        examples["F2"] = []

    # ============================================================
    # THEME G — Statistical anomalies
    # ============================================================
    print("\n=== THEME G — Statistical anomalies ===")

    # G1: Price outliers per Instrument (IQR-based)
    valid_px = df[df["Price_num"].notna() & (df["Price_num"] > 0) & (df["Instrument"] != "")]
    q1 = valid_px.groupby("Instrument")["Price_num"].quantile(0.25)
    q3 = valid_px.groupby("Instrument")["Price_num"].quantile(0.75)
    iqr = (q3 - q1)
    lo = (q1 - 3 * iqr)
    hi = (q3 + 3 * iqr)
    valid_px = valid_px.assign(
        _lo=valid_px["Instrument"].map(lo),
        _hi=valid_px["Instrument"].map(hi),
    )
    outlier_mask = (valid_px["Price_num"] < valid_px["_lo"]) | (valid_px["Price_num"] > valid_px["_hi"])
    add_finding(findings, "G1", theme="G. Statistical",
                name="Price outliers (per-instrument IQR fence, 3x)",
                rationale="A surveillance system needs to flag prices that don't make sense given the rest of the same-instrument activity (fat-finger; off-market prints; bad market-data). The 3*IQR fence is a conservative non-parametric outlier definition.",
                detection="For each Instrument with valid Price>0, compute Q1, Q3, IQR. Flag rows where Price ∉ [Q1-3·IQR, Q3+3·IQR].",
                count=int(outlier_mask.sum()),
                extras={"pct_of_priced_fills": round(100 * outlier_mask.sum() / max(len(valid_px), 1), 2)})
    examples["G1"] = (valid_px[outlier_mask].head(5)[["MessageId", "Instrument", "Price", "TotalVolume", "_lo", "_hi"]]
                       .astype(str).to_dict(orient="records"))

    # G2: Volume outliers per Instrument
    valid_vol = df[df["TotalVolume_num"].notna() & (df["TotalVolume_num"] > 0) & (df["Instrument"] != "")]
    vq1 = valid_vol.groupby("Instrument")["TotalVolume_num"].quantile(0.25)
    vq3 = valid_vol.groupby("Instrument")["TotalVolume_num"].quantile(0.75)
    viqr = (vq3 - vq1)
    vhi = (vq3 + 3 * viqr)
    valid_vol = valid_vol.assign(_vhi=valid_vol["Instrument"].map(vhi))
    voutlier_mask = valid_vol["TotalVolume_num"] > valid_vol["_vhi"]
    add_finding(findings, "G2", theme="G. Statistical",
                name="Volume outliers (per-instrument upper IQR fence, 3x)",
                rationale="A single block dramatically larger than the typical size for that instrument should be inspected — could be a legit block but might also be a fat-finger or marking-the-close trade.",
                detection="For each Instrument with valid Volume>0, flag rows where Volume > Q3 + 3·IQR.",
                count=int(voutlier_mask.sum()))
    examples["G2"] = (valid_vol[voutlier_mask].head(5)[["MessageId", "Instrument", "TotalVolume", "_vhi"]]
                       .astype(str).to_dict(orient="records"))

    # ============================================================
    # Persist
    # ============================================================
    FINDINGS.write_text(json.dumps(findings, indent=2, default=str), encoding="utf-8")
    EXAMPLES.write_text(json.dumps(examples, indent=2, default=str), encoding="utf-8")
    print(f"\nWrote: {FINDINGS.name}")
    print(f"Wrote: {EXAMPLES.name}")

    # Summary print
    print("\n=== P2 FINDINGS SUMMARY ===")
    by_theme: dict[str, list[tuple[str, str, int]]] = {}
    for key, info in findings.items():
        by_theme.setdefault(info["theme"], []).append((key, info["name"], info["count"]))
    for theme, items in sorted(by_theme.items()):
        print(f"\n{theme}")
        for k, name, c in items:
            print(f"  {k:<5} {name[:80]:<80} {c:>10,}")


if __name__ == "__main__":
    main()
