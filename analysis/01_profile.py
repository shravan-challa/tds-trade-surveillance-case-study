"""
Stage 1 — Raw profile of the synthetic trade surveillance CSV.

Goal: understand what's actually in the file BEFORE any cleaning. Read everything
as strings (no type coercion) so we see the data as the vendor ingestion would
first encounter it. Surface anything that smells like a defect — but only count
and describe at this stage; no fixes yet.
"""

from __future__ import annotations

import json
import sys
from collections import Counter
from pathlib import Path

import pandas as pd

DATA = Path(__file__).resolve().parents[1] / "data" / "synthetic_trade_data.csv"
OUT = Path(__file__).resolve().parent / "01_profile_output.json"

# Tokens that often mean "this column was meant to be null" but bypass real null detection
SUSPECT_NULL_TOKENS = {"", "NaN", "nan", "NULL", "null", "None", "none", "N/A", "n/a", "NA", "-", "?"}


def main() -> None:
    print(f"Reading {DATA} ...", flush=True)

    # Raw line count — USE UNIVERSAL NEWLINE MODE (text mode with newline=None).
    # The file contains a mix of \r\n and bare \r line terminators (4,193 bare CRs).
    # wc -l / binary-mode counting undercounts by exactly that many. Power Query,
    # Excel, and most modern CSV readers correctly use universal newlines, so this
    # is the right baseline for "rows the vendor would attempt to ingest".
    with open(DATA, "r", encoding="utf-8", newline=None) as f:
        raw_line_count = sum(1 for _ in f)
    raw_data_lines = raw_line_count - 1  # minus header
    # Also record byte-level newline composition for the audit trail
    raw_bytes = DATA.read_bytes()
    n_lf = raw_bytes.count(b"\n")
    n_cr = raw_bytes.count(b"\r")
    n_crlf = raw_bytes.count(b"\r\n")
    n_bare_cr = n_cr - n_crlf
    print(f"Raw file: {raw_line_count:,} lines ({raw_data_lines:,} data lines, universal-newline)", flush=True)
    print(f"Line endings: {n_crlf:,} \\r\\n  +  {n_bare_cr:,} bare \\r  +  {n_lf - n_crlf:,} bare \\n", flush=True)

    # Read with tolerant parser. We want to know: how many lines does pandas successfully
    # tokenize into 18 fields? The gap between raw_data_lines and parsed rows = rows the
    # vendor parser would either drop or silently merge.
    df = pd.read_csv(
        DATA,
        dtype=str,
        keep_default_na=False,
        na_filter=False,
        encoding="utf-8",
        engine="python",
        on_bad_lines="skip",
    )

    n_rows = len(df)
    csv_malformed = raw_data_lines - n_rows
    print(f"Parsed: {n_rows:,} rows x {len(df.columns)} cols", flush=True)
    print(f"CSV-malformed (skipped or merged): {csv_malformed:,} rows ({100*csv_malformed/raw_data_lines:.2f}%)", flush=True)

    # Header hygiene — column names might have whitespace too
    raw_columns = list(df.columns)
    stripped_columns = [c.strip() for c in raw_columns]
    header_whitespace_cols = [(r, s) for r, s in zip(raw_columns, stripped_columns) if r != s]

    profile: dict = {
        "raw_data_lines": raw_data_lines,
        "raw_line_endings": {
            "crlf_pairs": n_crlf,
            "bare_cr": n_bare_cr,
            "bare_lf": n_lf - n_crlf,
            "total_lf": n_lf,
            "total_cr": n_cr,
        },
        "parsed_row_count": n_rows,
        "csv_malformed_count": csv_malformed,
        "csv_malformed_pct": round(100 * csv_malformed / raw_data_lines, 4),
        "column_count": len(df.columns),
        "raw_columns": raw_columns,
        "stripped_columns": stripped_columns,
        "header_has_whitespace": len(header_whitespace_cols) > 0,
        "header_whitespace_examples": header_whitespace_cols[:5],
        "columns": {},
    }

    for raw_col in df.columns:
        col = raw_col  # keep the raw name for indexing
        series = df[col]

        # Basic stats
        n_null_pandas = 0  # we set na_filter=False, so this is always 0 — useful sanity check
        n_empty = int((series == "").sum())
        n_whitespace_only = int(series.str.fullmatch(r"\s+").fillna(False).sum())
        n_has_leading_or_trailing_ws = int(((series.str.len() > 0) & (series != series.str.strip())).sum())

        # Suspect null tokens (after strip, since trailing whitespace is its own defect)
        stripped = series.str.strip()
        suspect_token_counts = Counter()
        for token in SUSPECT_NULL_TOKENS:
            c = int((stripped == token).sum())
            if c > 0:
                suspect_token_counts[token] = c

        # Cardinality + sample
        non_empty_stripped = stripped[stripped != ""]
        nunique = int(non_empty_stripped.nunique())
        # Top-10 value counts on stripped values, excluding empty
        top_values = (
            non_empty_stripped.value_counts(dropna=False).head(10).to_dict()
        )
        # Length distribution (on stripped) — catches truncation / padding
        lengths = non_empty_stripped.str.len()
        len_stats = {
            "min": int(lengths.min()) if len(lengths) else None,
            "max": int(lengths.max()) if len(lengths) else None,
            "mean": round(float(lengths.mean()), 2) if len(lengths) else None,
        }

        # Sample 3 raw values (with whitespace preserved) to eyeball
        sample_raw = series.head(3).tolist()

        profile["columns"][col.strip()] = {
            "raw_header": raw_col,
            "n_empty": n_empty,
            "pct_empty": round(100 * n_empty / n_rows, 3),
            "n_whitespace_only": n_whitespace_only,
            "n_leading_or_trailing_whitespace": n_has_leading_or_trailing_ws,
            "pct_leading_or_trailing_whitespace": round(100 * n_has_leading_or_trailing_ws / n_rows, 3),
            "suspect_null_tokens": dict(suspect_token_counts),
            "nunique_non_empty_stripped": nunique,
            "stripped_length_stats": len_stats,
            "top_10_values_stripped": {str(k): int(v) for k, v in top_values.items()},
            "sample_raw_values": sample_raw,
        }

    OUT.write_text(json.dumps(profile, indent=2, default=str), encoding="utf-8")
    print(f"Wrote profile to {OUT}", flush=True)

    # Print a compact summary to stdout — high-level signals only
    print("\n=== HIGH-LEVEL SIGNALS ===")
    print(f"Rows: {n_rows:,}  Cols: {len(df.columns)}")
    if profile["header_has_whitespace"]:
        print(f"!! Header column names contain whitespace (e.g., {profile['header_whitespace_examples'][:3]})")
    print()
    print(f"{'Column':<22} {'%empty':>8} {'%ws_trim':>10} {'nunique':>10}  null_tokens")
    for cstrip, info in profile["columns"].items():
        tokens = ",".join(f"{k}={v}" for k, v in info["suspect_null_tokens"].items()) or "-"
        print(
            f"{cstrip:<22} {info['pct_empty']:>8.2f} {info['pct_leading_or_trailing_whitespace']:>10.2f} "
            f"{info['nunique_non_empty_stripped']:>10}  {tokens}"
        )


if __name__ == "__main__":
    main()
