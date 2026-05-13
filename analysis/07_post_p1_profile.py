"""
Stage P1 — Post-fix profile.

Profile the P1-cleaned dataset using the same dimensions as 01_profile.py so we
can produce a before/after comparison for the doc.
"""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
SRC = HERE / "cleaned_p1_final.tsv"
FLAGS_SRC = HERE / "06_row_flags.tsv"
STAGE1 = HERE / "01_profile_output.json"
OUT = HERE / "07_post_p1_profile.json"


def main() -> None:
    print(f"Reading {SRC.name} ...")
    df = pd.read_csv(SRC, sep="\t", dtype=str, keep_default_na=False, na_filter=False)
    flags = pd.read_csv(FLAGS_SRC, sep="\t")
    n = len(df)
    stage1 = json.loads(STAGE1.read_text(encoding="utf-8"))

    print(f"Rows: {n:,}")

    summary: dict = {
        "stage1_parsed_rows": stage1["parsed_row_count"],
        "stage1_raw_data_lines": stage1["raw_data_lines"],
        "stage1_csv_malformed": stage1["csv_malformed_count"],
        "p1_rows": n,
        "p1_recovered_vs_stage1": n - stage1["parsed_row_count"],
        "p1_parse_rate_vs_raw_pct": round(100 * n / stage1["raw_data_lines"], 2),
        "columns": {},
    }

    for col in df.columns:
        s = df[col]
        n_null = int((s == "").sum())
        pct_null = round(100 * n_null / n, 3)
        non_null = s[s != ""]
        nunique = int(non_null.nunique())

        # Stage-1 comparable stats (Stage 1 was raw with whitespace, so values are roughly comparable
        # after stripping; we already stripped during repair)
        s1 = stage1["columns"][col]
        s1_pct_ws = s1["pct_leading_or_trailing_whitespace"]
        s1_nunique = s1["nunique_non_empty_stripped"]
        s1_n_empty = s1["n_empty"]
        s1_pct_empty = s1["pct_empty"]
        s1_suspect_tokens_total = sum(s1.get("suspect_null_tokens", {}).values())

        summary["columns"][col] = {
            "stage1": {
                "pct_empty": s1_pct_empty,
                "pct_whitespace_padding": s1_pct_ws,
                "suspect_null_tokens_total": s1_suspect_tokens_total,
                "nunique": s1_nunique,
            },
            "p1": {
                "pct_null": pct_null,
                "n_null": n_null,
                "nunique": nunique,
            },
            "delta": {
                "pct_whitespace_padding_change": round(0 - s1_pct_ws, 2),  # always 0% after strip
                "nunique_change": nunique - s1_nunique,
            },
        }

    # Row-health: how many rows have ANY flag fired
    flag_cols = [c for c in flags.columns if c.startswith("FLAG_")]
    any_flag = (flags[flag_cols].sum(axis=1) > 0).sum()
    summary["row_health"] = {
        "total_rows": n,
        "rows_with_any_flag": int(any_flag),
        "pct_rows_with_any_flag": round(100 * int(any_flag) / n, 2),
        "rows_clean": int(n - any_flag),
        "pct_rows_clean": round(100 * (n - any_flag) / n, 2),
        "flag_counts": {c: int(flags[c].sum()) for c in flag_cols},
    }

    OUT.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print(f"Wrote: {OUT.name}")

    # Print compact summary
    print()
    print(f"Raw input lines (Stage 1): {stage1['raw_data_lines']:,}")
    print(f"Stage-1 strict-parsed:     {stage1['parsed_row_count']:,}")
    print(f"P1 rows:                   {n:,}")
    print(f"P1 recovered vs Stage-1:   {n - stage1['parsed_row_count']:+,}")
    print(f"P1 parse-rate of raw:      {100 * n / stage1['raw_data_lines']:.2f}%")
    print(f"Rows fully clean (no flag fired): {n - any_flag:,} ({100*(n-any_flag)/n:.2f}%)")
    print(f"Rows with at least one P1 flag:   {any_flag:,} ({100*any_flag/n:.2f}%)")


if __name__ == "__main__":
    main()
