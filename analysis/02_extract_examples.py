"""
Stage 1.5 — Extract concrete row-level examples for each defect class
identified in 01_profile.py. The Stage 1 documentation needs evidence,
not just counts.

For each defect we record up to 3 example rows: their parsed values and,
where useful, the raw line text. Output: examples.json + raw_bad_lines.json
"""

from __future__ import annotations

import csv
import io
import json
import re
from pathlib import Path

import pandas as pd

DATA = Path(__file__).resolve().parents[1] / "data" / "synthetic_trade_data.csv"
OUT = Path(__file__).resolve().parent / "02_examples.json"
RAW_BAD = Path(__file__).resolve().parent / "02_raw_bad_lines.json"

EXPECTED_COLS = 18


def find_raw_bad_lines(path: Path, expected_cols: int, max_examples: int = 10):
    """
    Walk the raw file line-by-line with the standard csv module and record lines
    that don't parse to expected_cols fields. Also detect lines with an odd number
    of unescaped quote characters (likely unclosed quote).
    """
    bad_lines = []
    quote_imbalance_lines = []

    with open(path, "r", encoding="utf-8", newline="") as f:
        reader = csv.reader(f)
        for lineno, row in enumerate(reader, start=1):
            if lineno == 1:
                continue  # header
            if len(row) != expected_cols:
                if len(bad_lines) < max_examples:
                    bad_lines.append({
                        "line_number": lineno,
                        "field_count": len(row),
                        "first_500_chars": (",".join(row))[:500],
                    })

    # Second pass — count quote-imbalance lines (likely root cause)
    with open(path, "r", encoding="utf-8") as f:
        for lineno, raw_line in enumerate(f, start=1):
            if lineno == 1:
                continue
            # Count unescaped double quotes
            quote_count = raw_line.count('"')
            if quote_count % 2 == 1:
                if len(quote_imbalance_lines) < max_examples:
                    quote_imbalance_lines.append({
                        "line_number": lineno,
                        "quote_count": quote_count,
                        "line_text": raw_line.rstrip("\n")[:400],
                    })

    return bad_lines, quote_imbalance_lines


def main() -> None:
    print("=== Finding raw-file CSV malformations ===", flush=True)
    bad_lines, odd_quote_lines = find_raw_bad_lines(DATA, EXPECTED_COLS)
    print(f"Bad-field-count examples collected: {len(bad_lines)}")
    print(f"Odd-quote-count examples collected: {len(odd_quote_lines)}")

    # Now load parsed data for cleanish examples
    df = pd.read_csv(
        DATA,
        dtype=str,
        keep_default_na=False,
        na_filter=False,
        encoding="utf-8",
        engine="python",
        on_bad_lines="skip",
    )
    # Standardise column names internally for searching, keep originals
    raw_cols = list(df.columns)
    df.columns = [c.strip() for c in raw_cols]

    def sample(mask, k=3):
        idx = df.index[mask][:k].tolist()
        rows = []
        for i in idx:
            rec = {c: df.at[i, c] for c in df.columns}
            rec["__row_index"] = int(i)
            rows.append(rec)
        return rows

    examples: dict = {
        "raw_bad_field_count_examples": bad_lines,
        "raw_odd_quote_examples": odd_quote_lines,
        "parsed_examples": {},
    }

    # Strip helpers
    def s(col):
        return df[col].str.strip()

    # 1. NULL literal in ExchangeId
    examples["parsed_examples"]["exchangeid_literal_NULL"] = sample(s("ExchangeId") == "NULL")
    examples["parsed_examples"]["exchangeid_OO_quote"] = sample(s("ExchangeId") == 'OO"')
    examples["parsed_examples"]["exchangeid_NONE"] = sample(s("ExchangeId") == "NONE")
    # Extremely long ExchangeId (quote contamination)
    examples["parsed_examples"]["exchangeid_extreme_length"] = sample(s("ExchangeId").str.len() > 100)

    # 2. MessageType case variants
    examples["parsed_examples"]["messagetype_uppercase"] = sample(s("MessageType") == "NEW ORDER")
    examples["parsed_examples"]["messagetype_lowercase"] = sample(s("MessageType") == "new order")
    examples["parsed_examples"]["messagetype_nan_literal"] = sample(s("MessageType") == "nan")
    examples["parsed_examples"]["messagetype_is_isin"] = sample(s("MessageType").str.match(r"^US[A-Z0-9]{10}$").fillna(False))

    # 3. TransactionTime — non-ISO values
    iso_re = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d+Z$")
    examples["parsed_examples"]["transactiontime_not_iso"] = sample(~s("TransactionTime").str.match(iso_re).fillna(False) & (s("TransactionTime") != ""))

    # 4. MessageDate format variants
    examples["parsed_examples"]["messagedate_dash_eu"] = sample(s("MessageDate") == "05-02-2026")
    examples["parsed_examples"]["messagedate_slash_us"] = sample(s("MessageDate") == "05/02/2026")
    examples["parsed_examples"]["messagedate_compact"] = sample(s("MessageDate") == "20260205")
    examples["parsed_examples"]["messagedate_numeric_leak"] = sample(s("MessageDate") == "0.0")

    # 5. MessageId — primary key issues
    examples["parsed_examples"]["messageid_empty"] = sample(s("MessageId") == "")
    examples["parsed_examples"]["messageid_NULL_token"] = sample(s("MessageId") == "NULL")
    examples["parsed_examples"]["messageid_NaN_token"] = sample(s("MessageId") == "NaN")
    examples["parsed_examples"]["messageid_numeric"] = sample(s("MessageId").str.match(r"^\d{1,4}$").fillna(False))

    # 6. Instrument bare-quote
    examples["parsed_examples"]["instrument_bare_quote"] = sample(s("Instrument") == '"')
    examples["parsed_examples"]["instrument_nan_literal"] = sample(s("Instrument") == "NaN")

    # 7. ISIN issues
    examples["parsed_examples"]["isin_too_short"] = sample(s("ISIN").str.len().between(1, 11) & (s("ISIN") != ""))
    examples["parsed_examples"]["isin_SYS_OMEGA_leak"] = sample(s("ISIN") == "SYS_OMEGA")

    # 8. BuyOrSell extra values
    examples["parsed_examples"]["buyorsell_USD"] = sample(s("BuyOrSell") == "USD")
    examples["parsed_examples"]["buyorsell_SYS_OMEGA"] = sample(s("BuyOrSell") == "SYS_OMEGA")
    examples["parsed_examples"]["buyorsell_nan"] = sample(s("BuyOrSell") == "nan")

    # 9. Trader NaN literal
    examples["parsed_examples"]["trader_NaN_literal"] = sample(s("Trader") == "NaN")
    examples["parsed_examples"]["trader_empty"] = sample(s("Trader") == "")

    # 10. Account == CounterPartyFirm collisions
    examples["parsed_examples"]["account_equals_counterparty"] = sample(s("Account") == s("CounterPartyFirm"))

    # 11. Column-shift signature row — multiple atypical values
    shift_mask = (s("BuyOrSell") == "USD") | (s("MessageType").str.match(r"^US[A-Z0-9]{10}$").fillna(False))
    examples["parsed_examples"]["column_shift_signatures"] = sample(shift_mask, k=5)

    OUT.write_text(json.dumps(examples, indent=2, default=str), encoding="utf-8")
    RAW_BAD.write_text(json.dumps({
        "bad_field_count_examples": bad_lines,
        "odd_quote_examples": odd_quote_lines,
    }, indent=2), encoding="utf-8")

    print(f"\nWrote: {OUT.name}")
    print(f"Wrote: {RAW_BAD.name}")
    print("\n=== EXAMPLE COUNTS PER CATEGORY ===")
    for k, v in examples["parsed_examples"].items():
        print(f"  {k:<40} {len(v)} examples")


if __name__ == "__main__":
    main()
