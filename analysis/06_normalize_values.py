"""
Stage P1 — Value-level normalisation, applied to the structurally-repaired
output of 05_repair_structural.py.

Each rule is a small named transformation with:
  - name (stable id used in docs)
  - column (target field)
  - rationale (the WHY)
  - action (the WHAT)
  - count of rows affected (logged after the transformation runs)

The pipeline is deliberately conservative: defects are converted to nulls and
flagged on a per-row basis rather than guessed-at. Aggressive recovery (e.g.,
column-shift re-alignment) is deferred to Stage P2 once a producer review of
the shift signature has been completed.

Outputs:
  ../analysis/cleaned_p1_final.tsv     — strict-typed, normalised data, TSV
  ../analysis/06_normalize_stats.json  — per-rule before/after counts (for the doc)
  ../analysis/06_row_flags.tsv         — per-row flag bitset (which rules fired on each row)
"""

from __future__ import annotations

import csv
import json
import re
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
SRC = HERE / "cleaned_stage_p1.tsv"
OUT_TSV = HERE / "cleaned_p1_final.tsv"
STATS = HERE / "06_normalize_stats.json"
FLAGS = HERE / "06_row_flags.tsv"


# Null-token universe — applied AFTER strip
GENERIC_NULL_TOKENS = {"", "NaN", "nan", "NaT", "NULL", "null", "None", "none", "N/A", "n/a", "NA"}

# Per-column extra null tokens (seeded placeholders)
EXTRA_NULLS = {
    "ExchangeId": {"NULL", "NONE", "ZZZZ", 'OO"'},
    "Instrument": {'"'},
}

CANONICAL_MESSAGE_TYPES = {"New Order", "Cancel Order", "Replace Order", "Fill", "Reject"}

MESSAGEID_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}\.RT\d+\.[A-Z]\.\d+$")
ISIN_PATTERN = re.compile(r"^[A-Z]{2}[A-Z0-9]{9}\d$")  # ISO 6166: 2-letter country + 9 alnum + check digit
ISO_TS_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$")
ACCOUNT_PATTERN = re.compile(r"^[A-Z]{3,5}\d{1,3}\.[A-Z]+$")
TRADER_PATTERN = re.compile(r"^TRD_\d{1,3}$")
EXCHANGEID_PATTERN = re.compile(r"^[A-Z]{2,6}$")

# Synthetic dataset canonical date (per dataset author):
# every record is intended to be on 2026-02-05 — alternate formats are
# corruption of that same date.
CANONICAL_DATE_STR = "2026-02-05"
DATE_FORMATS = [
    "%Y-%m-%d",  # 2026-02-05 (canonical)
    "%d-%m-%Y",  # 05-02-2026
    "%d-%m-%y",  # 05-02-26
    "%d/%m/%Y",  # 05/02/2026
    "%m/%d/%Y",  # 02/05/2026 (US — same outcome for 2026-02-05)
    "%Y%m%d",    # 20260205
]


def main() -> None:
    print(f"Reading {SRC.name} ...")
    df = pd.read_csv(SRC, sep="\t", dtype=str, keep_default_na=False, na_filter=False)
    n = len(df)
    print(f"Loaded {n:,} rows x {len(df.columns)} cols")

    stats: dict = {"input_rows": n, "rules": []}
    flags = pd.DataFrame(index=df.index)  # per-row flag bitset

    def log(rule_id: str, column: str, rationale: str, action: str, affected: int) -> None:
        stats["rules"].append({
            "rule_id": rule_id,
            "column": column,
            "rationale": rationale,
            "action": action,
            "rows_affected": int(affected),
        })
        print(f"  [{rule_id}] {column}: {affected:,} rows — {action}")

    # ---------- Rule N1 — Generic null-token unification ----------
    print("\n--- Null-token unification ---")
    for col in df.columns:
        is_generic_null = df[col].isin(GENERIC_NULL_TOKENS)
        extra = EXTRA_NULLS.get(col, set())
        is_extra_null = df[col].isin(extra) if extra else pd.Series(False, index=df.index)
        mask = is_generic_null | is_extra_null
        if mask.any():
            df.loc[mask, col] = ""  # uniform empty marker, converted to NaN at end
            flags[f"FLAG_NULL_{col}"] = mask.astype(int)
            extra_note = f" (plus {sorted(extra)} for this field)" if extra else ""
            log(
                f"N1.{col}",
                col,
                f"Multiple null tokens for the same logical state ({sorted(GENERIC_NULL_TOKENS)}{extra_note}) confuses every downstream null-check and creates spurious duplicates on PK fields.",
                f"Coerce {sorted(GENERIC_NULL_TOKENS) + sorted(extra)} -> null",
                int(mask.sum()),
            )

    # ---------- Rule N2 — MessageType case normalisation ----------
    print("\n--- MessageType case normalisation ---")
    mt = df["MessageType"].str.strip()
    # Title-case (each word's first letter capital, rest lower)
    titlecased = mt.str.title()
    changed = (titlecased != mt) & (mt != "")
    df["MessageType"] = titlecased.where(mt != "", "")
    flags["FLAG_MT_CASE"] = changed.astype(int)
    log(
        "N2",
        "MessageType",
        "Three case-variants of each enum value present ('New Order' / 'NEW ORDER' / 'new order'). Any case-sensitive downstream join, group-by, or surveillance rule will treat them as distinct event types.",
        "Coerce to Title Case (canonical form).",
        int(changed.sum()),
    )

    # Flag rows whose MessageType doesn't match the canonical enum after normalisation
    not_in_enum = (df["MessageType"] != "") & ~df["MessageType"].isin(CANONICAL_MESSAGE_TYPES)
    flags["FLAG_MT_UNKNOWN_ENUM"] = not_in_enum.astype(int)
    log(
        "N2b",
        "MessageType",
        "Values that don't match the canonical enum after case normalisation are almost certainly column-shift contamination (e.g., an ISIN sitting in the MessageType column).",
        f"Flag rows where MessageType ∉ {sorted(CANONICAL_MESSAGE_TYPES)} (no value change — Stage P2 will decide on recovery vs quarantine).",
        int(not_in_enum.sum()),
    )

    # ---------- Rule N3 — MessageDate canonicalisation ----------
    print("\n--- MessageDate canonicalisation ---")
    raw_dates = df["MessageDate"].astype(str)
    parsed = pd.Series(pd.NaT, index=df.index, dtype="object")
    fmt_counts: dict[str, int] = {}
    for fmt in DATE_FORMATS:
        unparsed_mask = parsed.isna() & (raw_dates != "")
        if not unparsed_mask.any():
            break
        attempt = pd.to_datetime(raw_dates[unparsed_mask], format=fmt, errors="coerce")
        ok = attempt.notna()
        parsed.loc[attempt[ok].index] = attempt[ok]
        fmt_counts[fmt] = int(ok.sum())

    canonical = parsed.apply(lambda v: v.strftime("%Y-%m-%d") if pd.notna(v) else "")
    changed_dates = (canonical != raw_dates) & (raw_dates != "") & (canonical != "")
    unparsable = (raw_dates != "") & (canonical == "")
    df["MessageDate"] = canonical
    flags["FLAG_DATE_REFORMATTED"] = changed_dates.astype(int)
    flags["FLAG_DATE_UNPARSABLE"] = unparsable.astype(int)

    log(
        "N3a",
        "MessageDate",
        "≥6 distinct date formats present in this column, including the ambiguous DD-MM vs MM-DD case. Free-form date strings break downstream date logic and surveillance windowing.",
        f"Parse against an explicit format list {DATE_FORMATS} and re-emit as ISO YYYY-MM-DD. Per-format match counts: " + json.dumps(fmt_counts),
        int(changed_dates.sum()),
    )
    log(
        "N3b",
        "MessageDate",
        "Values that match NO known format (e.g., '0.0', '12.36') are column-shift artifacts where a Price value has shifted into the date column.",
        "Set unparsable values to null and flag the row.",
        int(unparsable.sum()),
    )

    # ---------- Rule N4 — TransactionTime ISO validation ----------
    print("\n--- TransactionTime validation ---")
    tt = df["TransactionTime"]
    is_iso = tt.str.match(ISO_TS_PATTERN).fillna(False)
    bad_tt = (tt != "") & ~is_iso
    flags["FLAG_TT_NOT_ISO"] = bad_tt.astype(int)
    df.loc[bad_tt, "TransactionTime"] = ""  # null out — these are 'Buy'/'Sell'/'nan' contaminants
    log(
        "N4",
        "TransactionTime",
        "Timestamps must be ISO-8601 with millisecond precision and Z (UTC). Non-conforming values are column-shift artifacts (top examples: 'Sell', 'Buy', 'nan').",
        "Set non-ISO timestamps to null and flag the row.",
        int(bad_tt.sum()),
    )

    # ---------- Rule N5 — MessageId primary-key validation ----------
    print("\n--- MessageId primary-key validation ---")
    mid = df["MessageId"]
    is_valid_pk = mid.str.match(MESSAGEID_PATTERN).fillna(False)
    bad_pk = (mid != "") & ~is_valid_pk
    null_pk = mid == ""
    flags["FLAG_PK_INVALID"] = bad_pk.astype(int)
    flags["FLAG_PK_NULL"] = null_pk.astype(int)
    df.loc[bad_pk, "MessageId"] = ""
    log(
        "N5a",
        "MessageId",
        "Primary keys must conform to the YYYY-MM-DD.RTnnn.P.nnn pattern. Values like '100', '400' (TotalVolume) and '1700' (TotalVolume) are column-shift contamination.",
        "Set malformed PK values to null and flag the row (the row remains in the dataset but is now correctly identified as having a missing PK).",
        int(bad_pk.sum()),
    )
    log(
        "N5b",
        "MessageId",
        "Primary keys cannot be null — a record with no PK cannot be deduplicated, reconciled to upstream OMS, or referenced by downstream Cancel/Fill events.",
        "Flag rows with empty MessageId (kept in dataset but will be excluded from PK-dependent metrics in Stage P3).",
        int(null_pk.sum()),
    )

    # PK uniqueness
    non_null_pk = df.loc[df["MessageId"] != "", "MessageId"]
    dup_pks = non_null_pk[non_null_pk.duplicated(keep=False)]
    n_dup_pk = len(dup_pks)
    log(
        "N5c",
        "MessageId",
        "Duplicate PKs imply either (a) replay-from-upstream, (b) merge artifacts from the malformed-quote contamination, or (c) seed-time duplication. Each requires a different remediation.",
        f"Detected {n_dup_pk:,} rows participating in PK duplicates (covering {non_null_pk.duplicated().sum():,} duplicate occurrences). Flagged but not modified at P1.",
        n_dup_pk,
    )
    flags["FLAG_PK_DUPLICATE"] = df["MessageId"].isin(dup_pks).astype(int)

    # ---------- Rule N6 — ISIN format validation ----------
    print("\n--- ISIN validation ---")
    isin = df["ISIN"]
    bad_isin = (isin != "") & ~isin.str.match(ISIN_PATTERN).fillna(False)
    flags["FLAG_ISIN_INVALID"] = bad_isin.astype(int)
    df.loc[bad_isin, "ISIN"] = ""
    log(
        "N6",
        "ISIN",
        "ISIN is a 12-character ISO 6166 identifier (2-letter country + 9 alnum + check digit). Non-conforming values include 'SYS_OMEGA' (column shift from TransactionSource) and any sub-12-char strings.",
        "Set non-ISIN-shaped values to null and flag the row. NB: we validate format only, not the check-digit, in this stage.",
        int(bad_isin.sum()),
    )

    # ---------- Rule N7 — BuyOrSell enum validation ----------
    print("\n--- BuyOrSell enum validation ---")
    side = df["BuyOrSell"]
    valid_side = {"Buy", "Sell"}
    bad_side = (side != "") & ~side.isin(valid_side)
    flags["FLAG_SIDE_INVALID"] = bad_side.astype(int)
    df.loc[bad_side, "BuyOrSell"] = ""
    log(
        "N7",
        "BuyOrSell",
        f"Side is a closed enum: {sorted(valid_side)}. Observed contaminants include 'USD' (Currency) and 'SYS_OMEGA' (TransactionSource) — same shift signature as the other affected fields.",
        "Set non-{Buy,Sell} values to null and flag the row.",
        int(bad_side.sum()),
    )

    # ---------- Rule N8 — Price coercion ----------
    print("\n--- Price coercion ---")
    price_str = df["Price"]
    price_num = pd.to_numeric(price_str, errors="coerce")
    bad_price = (price_str != "") & price_num.isna()
    neg_price = price_num < 0
    flags["FLAG_PRICE_UNPARSABLE"] = bad_price.astype(int)
    flags["FLAG_PRICE_NEGATIVE"] = neg_price.fillna(False).astype(int)
    df["Price"] = price_num.where(price_num.notna() & (price_num >= 0), np.nan).astype(object).where(lambda s: s.notna(), "")
    log(
        "N8a",
        "Price",
        "Price must be a non-negative decimal. Non-parseable strings indicate contamination.",
        "Coerce to float; non-parseable -> null + flag.",
        int(bad_price.sum()),
    )
    log(
        "N8b",
        "Price",
        "Negative prices are nonsensical for surveillance — a negative price cannot trigger any meaningful spoofing/layering/wash-trade rule.",
        "Set negative prices to null + flag.",
        int(neg_price.fillna(False).sum()),
    )

    # ---------- Rule N9 — TotalVolume coercion ----------
    print("\n--- TotalVolume coercion ---")
    vol_str = df["TotalVolume"]
    vol_num = pd.to_numeric(vol_str, errors="coerce")
    bad_vol = (vol_str != "") & vol_num.isna()
    non_pos_vol = (vol_num <= 0).fillna(False)
    flags["FLAG_VOL_UNPARSABLE"] = bad_vol.astype(int)
    flags["FLAG_VOL_NONPOSITIVE"] = non_pos_vol.astype(int)
    df["TotalVolume"] = vol_num.where(vol_num.notna() & (vol_num > 0), np.nan).astype(object).where(lambda s: s.notna(), "")
    log(
        "N9a",
        "TotalVolume",
        "Volume must be a positive integer; zero or negative cannot produce a surveillance event.",
        "Coerce to numeric; non-parseable or ≤0 -> null + flag.",
        int(bad_vol.sum() + non_pos_vol.sum()),
    )

    # ---------- Rule N10 — Account / CounterPartyFirm / Trader / ExchangeId format checks ----------
    print("\n--- Identifier format checks (flag-only, no value change) ---")
    for col, pattern, name in [
        ("Account", ACCOUNT_PATTERN, "ACCOUNT"),
        ("CounterPartyFirm", ACCOUNT_PATTERN, "CPTY"),
        ("Trader", TRADER_PATTERN, "TRADER"),
        ("ExchangeId", EXCHANGEID_PATTERN, "EXCH"),
    ]:
        s = df[col]
        bad = (s != "") & ~s.str.match(pattern).fillna(False)
        flags[f"FLAG_{name}_FORMAT"] = bad.astype(int)
        # Special: ExchangeId — null out clearly bogus values that have already
        # been mapped to "" via the N1 step (placeholders NULL/NONE/ZZZZ etc.).
        # For Account/CPTY/Trader we ONLY flag, we do not null, because the data
        # may still be informative even if not strictly pattern-matching.
        log(
            f"N10.{name}",
            col,
            f"Format check against {pattern.pattern}. Non-matching values are likely contamination or schema drift; surfaced for review.",
            "Flag-only (no value change at P1).",
            int(bad.sum()),
        )

    # ---------- Final null promotion: empty strings -> NaN ----------
    for col in df.columns:
        df[col] = df[col].replace("", np.nan)

    # ---------- Cross-row column-shift score ----------
    print("\n--- Column-shift score ---")
    shift_indicators = [
        "FLAG_MT_UNKNOWN_ENUM",
        "FLAG_DATE_UNPARSABLE",
        "FLAG_TT_NOT_ISO",
        "FLAG_PK_INVALID",
        "FLAG_ISIN_INVALID",
        "FLAG_SIDE_INVALID",
    ]
    available = [c for c in shift_indicators if c in flags.columns]
    flags["COLUMN_SHIFT_SCORE"] = flags[available].sum(axis=1)
    n_shifted_2plus = int((flags["COLUMN_SHIFT_SCORE"] >= 2).sum())
    log(
        "N11",
        "(row-level)",
        "A row that fails ≥2 independent format checks is almost certainly column-shifted (single random defects are unlikely to coincide). This is the population that Stage P2 will either re-align or quarantine.",
        f"Compute COLUMN_SHIFT_SCORE = sum of {available}. Rows with score≥2 = {n_shifted_2plus:,}.",
        n_shifted_2plus,
    )

    # ---------- Persist outputs ----------
    df.to_csv(OUT_TSV, sep="\t", index=False)
    flags.to_csv(FLAGS, sep="\t", index=False)
    STATS.write_text(json.dumps(stats, indent=2), encoding="utf-8")

    print(f"\nWrote: {OUT_TSV.name} ({len(df):,} rows)")
    print(f"Wrote: {FLAGS.name}")
    print(f"Wrote: {STATS.name}")


if __name__ == "__main__":
    main()
