"""
Test suite for cleaned_p1_final.tsv — verifies that every invariant the
P1 pipeline (06_normalize_values.py) claims to enforce actually holds on
the output file, AND that the headline counts cited in the deliverables
(stage_p1_repair_report.docx, methodology_journal.docx) match the file.

Each test prints PASS / FAIL with the actual count vs the asserted
condition. Exit code is non-zero if ANY test fails.

The point: if a reviewer at TD asks "how do you know your P1 pipeline did
what you say it did?" — the answer is "I ran this script, here's the
output, here's the assertion that fired on every rule, every count is
either 0 (defect eliminated) or matches the number in the report."
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
SRC = HERE / "cleaned_p1_final.tsv"
POST_PROFILE = HERE / "07_post_p1_profile.json"
NORMALIZE_STATS = HERE / "06_normalize_stats.json"

# -- Invariants the P1 pipeline claims to enforce ----------------------------
CANONICAL_MESSAGE_TYPES = {"New Order", "Cancel Order", "Replace Order", "Fill", "Reject"}
LEGACY_NULL_TOKENS = {"NaN", "nan", "NaT", "NULL", "null", "None", "none", "N/A", "n/a", "NA"}
EXCHANGEID_EXTRA_NULLS = {"NULL", "NONE", "ZZZZ", 'OO"'}

MESSAGEID_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}\.RT\d+\.[A-Z]\.\d+$")
ISIN_PATTERN = re.compile(r"^[A-Z]{2}[A-Z0-9]{9}\d$")
ISO_TS_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$")
DATE_ISO_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}$")

EXPECTED_COLUMNS = [
    "ExchangeId", "MessageType", "TransactionTime", "MessageDate",
    "MessageId", "LinkMessageId", "ParentOrderId", "Instrument",
    "ISIN", "BuyOrSell", "Price", "TotalVolume",
    "Account", "CounterPartyFirm", "Trader", "TransactionSource",
    "Currency", "Flags",
]

EXPECTED_ROW_COUNT = 242_429  # cited in stage_p1_repair_report.docx + methodology_journal §6.3

# Per-column null counts cited in 07_post_p1_profile.json — pulled at runtime
# so the test stays in sync with the JSON, not hard-coded.

results: list[tuple[str, str, str, str]] = []  # (id, name, status, detail)


def record(test_id: str, name: str, ok: bool, detail: str) -> None:
    status = "PASS" if ok else "FAIL"
    results.append((test_id, name, status, detail))
    marker = "[OK]  " if ok else "[FAIL]"
    print(f"  {marker} {test_id:8s} {name}")
    if detail:
        for line in detail.splitlines():
            print(f"             {line}")


def section(title: str) -> None:
    print(f"\n{title}")
    print("-" * len(title))


def main() -> int:
    print(f"Loading {SRC.name} ...")
    df = pd.read_csv(SRC, sep="\t", dtype=str, keep_default_na=True, na_values=[""])
    print(f"Loaded: {len(df):,} rows x {len(df.columns)} columns")

    profile = json.loads(POST_PROFILE.read_text(encoding="utf-8"))

    # ============================================================
    # T0 — schema & shape
    # ============================================================
    section("T0 — Schema and shape")

    record(
        "T0.1", "Column set & order matches expected schema",
        list(df.columns) == EXPECTED_COLUMNS,
        f"actual={list(df.columns)}",
    )
    record(
        "T0.2", f"Row count == {EXPECTED_ROW_COUNT:,} (matches stage_p1_repair_report)",
        len(df) == EXPECTED_ROW_COUNT,
        f"actual={len(df):,}",
    )

    # ============================================================
    # T1 — N1 null-token unification: no legacy null tokens survived
    # ============================================================
    section("T1 — N1 null-token unification")

    found_legacy = {}
    for col in df.columns:
        s = df[col].dropna().astype(str)
        hits = s[s.isin(LEGACY_NULL_TOKENS)]
        if len(hits) > 0:
            found_legacy[col] = int(len(hits))
    record(
        "T1.1", "Zero literal legacy null tokens (NaN/NULL/None/N/A/...) survived",
        len(found_legacy) == 0,
        f"violations={found_legacy}" if found_legacy else "0 across all columns",
    )

    exch = df["ExchangeId"].dropna().astype(str)
    exch_extra = exch[exch.isin(EXCHANGEID_EXTRA_NULLS)]
    record(
        "T1.2", "ExchangeId carries no NULL/NONE/ZZZZ/OO\" placeholders",
        len(exch_extra) == 0,
        f"violations={int(len(exch_extra))}",
    )

    # ============================================================
    # T2 — N2 MessageType
    # ============================================================
    section("T2 — N2 MessageType normalisation")

    mt_non_null = df["MessageType"].dropna().astype(str)
    bad_mt = mt_non_null[~mt_non_null.isin(CANONICAL_MESSAGE_TYPES)]
    record(
        "T2.1", "Every non-null MessageType in canonical enum {New Order, Cancel Order, Replace Order, Fill, Reject}",
        len(bad_mt) == 0,
        f"distinct values={sorted(mt_non_null.unique().tolist())}; violations={int(len(bad_mt))}",
    )
    nunique_mt = mt_non_null.nunique()
    expected_mt_unique = profile["columns"]["MessageType"]["p1"]["nunique"]
    record(
        "T2.2", f"nunique(MessageType) matches 07_post_p1_profile (={expected_mt_unique})",
        nunique_mt == expected_mt_unique,
        f"actual={nunique_mt}",
    )

    # ============================================================
    # T3 — N3 MessageDate canonicalisation
    # ============================================================
    section("T3 — N3 MessageDate canonicalisation")

    md = df["MessageDate"].dropna().astype(str)
    md_bad_format = md[~md.str.match(DATE_ISO_PATTERN)]
    record(
        "T3.1", "Every non-null MessageDate matches YYYY-MM-DD",
        len(md_bad_format) == 0,
        f"violations={int(len(md_bad_format))}",
    )

    md_parsed = pd.to_datetime(md, format="%Y-%m-%d", errors="coerce")
    md_unparsable = md_parsed.isna().sum()
    record(
        "T3.2", "Every non-null MessageDate parses as a real calendar date",
        md_unparsable == 0,
        f"unparsable={int(md_unparsable)}",
    )

    # ============================================================
    # T4 — N4 TransactionTime
    # ============================================================
    section("T4 — N4 TransactionTime ISO 8601")

    tt = df["TransactionTime"].dropna().astype(str)
    tt_bad = tt[~tt.str.match(ISO_TS_PATTERN)]
    record(
        "T4.1", "Every non-null TransactionTime matches ISO 8601 .NNNZ pattern",
        len(tt_bad) == 0,
        f"violations={int(len(tt_bad))}",
    )

    tt_parsed = pd.to_datetime(tt, format="%Y-%m-%dT%H:%M:%S.%fZ", errors="coerce", utc=True)
    record(
        "T4.2", "Every non-null TransactionTime is parseable to a tz-aware datetime",
        tt_parsed.isna().sum() == 0,
        f"unparsable={int(tt_parsed.isna().sum())}",
    )

    # ============================================================
    # T5 — N5 MessageId primary key
    # ============================================================
    section("T5 — N5 MessageId primary key")

    mid_non_null = df["MessageId"].dropna().astype(str)
    mid_bad_format = mid_non_null[~mid_non_null.str.match(MESSAGEID_PATTERN)]
    record(
        "T5.1", "Every non-null MessageId matches YYYY-MM-DD.RTnnn.X.nnn",
        len(mid_bad_format) == 0,
        f"violations={int(len(mid_bad_format))}",
    )

    expected_pk_null = profile["row_health"]["flag_counts"]["FLAG_PK_NULL"]
    actual_pk_null = int(df["MessageId"].isna().sum())
    record(
        "T5.2", f"MessageId null count matches 07_post_p1_profile (={expected_pk_null:,})",
        actual_pk_null == expected_pk_null,
        f"actual={actual_pk_null:,}",
    )

    expected_pk_dup = profile["row_health"]["flag_counts"]["FLAG_PK_DUPLICATE"]
    actual_pk_dup = int(df["MessageId"].dropna().duplicated(keep=False).sum())
    record(
        "T5.3", f"PK-duplicate row count matches 07_post_p1_profile (={expected_pk_dup:,})",
        actual_pk_dup == expected_pk_dup,
        f"actual={actual_pk_dup:,}",
    )

    # ============================================================
    # T6 — N6 ISIN
    # ============================================================
    section("T6 — N6 ISIN format")

    isin = df["ISIN"].dropna().astype(str)
    isin_wrong_len = isin[isin.str.len() != 12]
    record(
        "T6.1", "Every non-null ISIN is exactly 12 characters",
        len(isin_wrong_len) == 0,
        f"violations={int(len(isin_wrong_len))}",
    )
    isin_bad_format = isin[~isin.str.match(ISIN_PATTERN)]
    record(
        "T6.2", "Every non-null ISIN matches ISO 6166 format ^[A-Z]{2}[A-Z0-9]{9}\\d$",
        len(isin_bad_format) == 0,
        f"violations={int(len(isin_bad_format))}",
    )

    # ============================================================
    # T7 — N7 BuyOrSell
    # ============================================================
    section("T7 — N7 BuyOrSell enum")

    side = df["BuyOrSell"].dropna().astype(str)
    bad_side = side[~side.isin({"Buy", "Sell"})]
    record(
        "T7.1", "Every non-null BuyOrSell in {Buy, Sell}",
        len(bad_side) == 0,
        f"distinct values={sorted(side.unique().tolist())}; violations={int(len(bad_side))}",
    )

    # ============================================================
    # T8 — N8 Price
    # ============================================================
    section("T8 — N8 Price")

    price_str = df["Price"].dropna()
    price_num = pd.to_numeric(price_str, errors="coerce")
    record(
        "T8.1", "Every non-null Price is numeric",
        price_num.isna().sum() == 0,
        f"non-numeric survivors={int(price_num.isna().sum())}",
    )
    record(
        "T8.2", "No negative Price values",
        (price_num.dropna() < 0).sum() == 0,
        f"negative survivors={int((price_num.dropna() < 0).sum())}",
    )

    # ============================================================
    # T9 — N9 TotalVolume
    # ============================================================
    section("T9 — N9 TotalVolume")

    vol_str = df["TotalVolume"].dropna()
    vol_num = pd.to_numeric(vol_str, errors="coerce")
    record(
        "T9.1", "Every non-null TotalVolume is numeric",
        vol_num.isna().sum() == 0,
        f"non-numeric survivors={int(vol_num.isna().sum())}",
    )
    record(
        "T9.2", "Every non-null TotalVolume strictly > 0",
        (vol_num.dropna() <= 0).sum() == 0,
        f"non-positive survivors={int((vol_num.dropna() <= 0).sum())}",
    )

    # ============================================================
    # T10 — Cross-check per-column null counts against post-P1 profile JSON
    # ============================================================
    section("T10 — Per-column null counts match 07_post_p1_profile.json")

    for col in EXPECTED_COLUMNS:
        if col not in profile["columns"]:
            continue
        expected = profile["columns"][col]["p1"]["n_null"]
        actual = int(df[col].isna().sum())
        record(
            f"T10.{col[:5]}", f"{col}: null count matches profile (expected={expected:,})",
            actual == expected,
            f"actual={actual:,}",
        )

    # ============================================================
    # T11 — Surveillance-domain probes (sanity, not pipeline guarantees)
    # ============================================================
    section("T11 — Surveillance-domain sanity probes")

    # T11.1 — Currency is single-valued (USD) per Stage 1 finding
    cur_unique = df["Currency"].dropna().unique().tolist()
    record(
        "T11.1", "Currency column is single-valued (matches Stage-1 finding 'USD only')",
        cur_unique == ["USD"],
        f"actual={cur_unique}",
    )

    # T11.2 — TransactionSource is single-valued (SYS_OMEGA)
    src_unique = df["TransactionSource"].dropna().unique().tolist()
    record(
        "T11.2", "TransactionSource column is single-valued (SYS_OMEGA)",
        src_unique == ["SYS_OMEGA"],
        f"actual={src_unique}",
    )

    # T11.3 — Every MessageId resolves consistently if non-null + non-duplicate
    mid_clean = df.loc[df["MessageId"].notna(), "MessageId"]
    pct_unique = mid_clean.nunique() / len(mid_clean) * 100 if len(mid_clean) > 0 else 0
    record(
        "T11.3", ">=90% of non-null MessageIds are unique (sanity probe -- not a pipeline guarantee)",
        pct_unique >= 90.0,
        f"unique pct={pct_unique:.2f}% ({mid_clean.nunique():,} / {len(mid_clean):,})",
    )

    # T11.4 — Cross-check headline metric: 96.70% parse rate vs 250,714 raw lines
    raw_lines = profile["stage1_raw_data_lines"]
    parse_rate = len(df) / raw_lines * 100
    expected_parse_rate = profile["p1_parse_rate_vs_raw_pct"]
    record(
        "T11.4", f"Parse rate vs raw lines = {expected_parse_rate}% (headline P1 metric)",
        abs(parse_rate - expected_parse_rate) < 0.05,
        f"actual={parse_rate:.2f}% (= {len(df):,} / {raw_lines:,})",
    )

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
