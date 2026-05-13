"""
Stage P1 — Structural repair of the CSV.

Two problems prevent strict-CSV ingestion of 104,410 rows:
  1. Mixed line endings (CRLF + bare CR).
  2. Pervasive ' , ' (space-comma-space) field separators that defeat CSV quote
     recognition. RFC 4180 quoted fields only work when the opening quote is the
     FIRST character of the field — a leading space turns the quote into a literal
     character, so embedded newlines inside what was meant to be a quoted field
     become record separators instead.

Fix strategy:
  - Read raw bytes, normalise line endings to \\n only.
  - Use csv.reader with skipinitialspace=True so leading whitespace doesn't
    defeat quoting.
  - Strip leading/trailing whitespace from every parsed field.
  - Each row's field count is checked against the 18-field schema. Rows that
    parse to exactly 18 fields are written to the clean output; everything else
    is written to a quarantine file with the original line number(s) preserved.

Outputs:
  ../analysis/cleaned_stage_p1.tsv          — strict 18-field rows, tab-delimited, whitespace-stripped
  ../analysis/quarantine_stage_p1.tsv       — rows still malformed after repair (for inspection)
  ../analysis/05_repair_stats.json          — counts in / out / recovered
"""

from __future__ import annotations

import csv
import io
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
DATA = HERE.parent / "data" / "synthetic_trade_data.csv"
CLEAN_OUT = HERE / "cleaned_stage_p1.tsv"
QUAR_OUT = HERE / "quarantine_stage_p1.tsv"
STATS_OUT = HERE / "05_repair_stats.json"

EXPECTED_COLS = 18

# csv module has a default field size limit (~131k); a couple of contaminated
# rows in this file have ExchangeId values >46k chars. Raise the cap so we can
# read those rows in order to QUARANTINE them rather than blow up.
csv.field_size_limit(10_000_000)


def main() -> None:
    raw_bytes = DATA.read_bytes()
    print(f"Read {len(raw_bytes):,} bytes")

    # Normalise line endings: any \r not followed by \n becomes \n; \r\n becomes \n.
    # Do this on bytes for speed and certainty.
    normalised = raw_bytes.replace(b"\r\n", b"\n").replace(b"\r", b"\n")
    n_lines_after = normalised.count(b"\n")
    print(f"Lines after newline normalisation: {n_lines_after:,}")

    # Decode and feed to csv.reader. skipinitialspace=True is the key flag:
    # it lets ` "..."` be recognised as a quoted field rather than as a literal
    # leading space followed by a quote character.
    text = normalised.decode("utf-8")
    reader = csv.reader(io.StringIO(text), skipinitialspace=True)

    rows = []
    quarantine = []
    header_raw = None
    row_idx = 0
    for record in reader:
        # First non-empty record is the header
        if header_raw is None:
            header_raw = [c.strip() for c in record]
            print(f"Header ({len(header_raw)} cols): {header_raw}")
            continue
        row_idx += 1
        # Strip trailing whitespace from every field (skipinitialspace handles leading)
        stripped = [c.strip() for c in record]
        if len(stripped) == EXPECTED_COLS:
            rows.append(stripped)
        else:
            # Keep the first 400 chars of the joined raw row for inspection
            preview = ",".join(stripped)[:400]
            quarantine.append({
                "csv_record_index": row_idx,  # the Nth CSV record after the header
                "field_count": len(stripped),
                "first_400_chars": preview,
            })

    print(f"\nParsed records (post-repair): {row_idx:,}")
    print(f"  Clean 18-field rows: {len(rows):,}")
    print(f"  Quarantined: {len(quarantine):,}")

    # Baselines for comparison
    baseline_raw_data_lines = 250_714
    baseline_parsed = 146_304
    recovered = len(rows) - baseline_parsed
    print(f"\nRecovery vs Stage-1 baseline (146,304 strict-parsed):")
    print(f"  Newly recovered rows: {recovered:+,}")
    print(f"  Recovery rate of previously-malformed (104,410): {100*recovered/104_410:.2f}%")
    print(f"  Final parse rate: {100*len(rows)/baseline_raw_data_lines:.2f}% of {baseline_raw_data_lines:,} raw data lines")

    # Write clean output as TSV (no need for quoting; whitespace already stripped)
    with open(CLEAN_OUT, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter="\t", quoting=csv.QUOTE_NONE, escapechar="\\")
        w.writerow(header_raw)
        for r in rows:
            # Defensive: replace any embedded tab/newline that survived
            safe = [str(c).replace("\t", " ").replace("\n", " ").replace("\r", " ") for c in r]
            w.writerow(safe)

    with open(QUAR_OUT, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter="\t", quoting=csv.QUOTE_MINIMAL)
        w.writerow(["csv_record_index", "field_count", "first_400_chars"])
        for q in quarantine:
            w.writerow([q["csv_record_index"], q["field_count"], q["first_400_chars"]])

    stats = {
        "input_file_bytes": len(raw_bytes),
        "lines_after_eol_normalisation": n_lines_after,
        "baseline_raw_data_lines": baseline_raw_data_lines,
        "baseline_strict_parsed": baseline_parsed,
        "baseline_malformed": 104_410,
        "post_repair_records_attempted": row_idx,
        "post_repair_clean_rows": len(rows),
        "post_repair_quarantine_rows": len(quarantine),
        "recovered_rows_vs_baseline": recovered,
        "recovery_rate_of_malformed_pct": round(100 * recovered / 104_410, 2),
        "final_parse_rate_vs_raw_pct": round(100 * len(rows) / baseline_raw_data_lines, 2),
    }
    STATS_OUT.write_text(json.dumps(stats, indent=2), encoding="utf-8")
    print(f"\nWrote: {CLEAN_OUT.name}")
    print(f"Wrote: {QUAR_OUT.name}")
    print(f"Wrote: {STATS_OUT.name}")


if __name__ == "__main__":
    main()
