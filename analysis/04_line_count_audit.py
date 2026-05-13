"""
Line-count audit — diagnose the discrepancy between:
  - wc -l (Git Bash):            246,522 lines (incl. header)
  - Power Query reported:        250,715 rows (incl. header)
  - pandas python engine parsed: 146,304 rows + skipped malformed

Possible causes:
  (a) Line-ending mix: bare CR (\r) is a logical line break in Power Query's
      universal-newline reader but is NOT counted by Git Bash's wc -l (which
      counts \n characters only).
  (b) Unicode line separators (U+2028, U+2029, U+0085) treated as line breaks
      by some readers.
  (c) Quote interpretation differences (CR/LF inside open quotes).
  (d) A NUL or other control char that confuses one reader.

This script counts each of these independently against the raw bytes so we
can attribute the gap exactly.
"""

from __future__ import annotations

import io
from pathlib import Path

DATA = Path(__file__).resolve().parents[1] / "data" / "synthetic_trade_data.csv"


def main() -> None:
    raw = DATA.read_bytes()
    size = len(raw)
    print(f"File size: {size:,} bytes  ({size/1024/1024:.2f} MiB)")

    # BOM check
    bom_utf8 = raw[:3] == b"\xef\xbb\xbf"
    bom_utf16_le = raw[:2] == b"\xff\xfe"
    bom_utf16_be = raw[:2] == b"\xfe\xff"
    print(f"BOM: utf-8={bom_utf8}  utf-16-le={bom_utf16_le}  utf-16-be={bom_utf16_be}")

    # Raw byte-level newline counts
    n_lf = raw.count(b"\n")
    n_cr = raw.count(b"\r")
    n_crlf = raw.count(b"\r\n")
    n_bare_cr = n_cr - n_crlf
    n_bare_lf = n_lf - n_crlf
    print(f"\nByte-level newline counts:")
    print(f"  \\n  total              : {n_lf:,}")
    print(f"  \\r  total              : {n_cr:,}")
    print(f"  \\r\\n pairs             : {n_crlf:,}")
    print(f"  bare \\r  (= \\r - \\r\\n) : {n_bare_cr:,}")
    print(f"  bare \\n  (= \\n - \\r\\n) : {n_bare_lf:,}")

    # Universal-newline line count (Python text mode treats \r, \n, \r\n equally)
    # This is what most modern CSV readers do (Power Query included for CSV).
    with open(DATA, "r", encoding="utf-8", newline=None) as f:
        univ_lines = sum(1 for _ in f)
    print(f"\nUniversal-newline line count (Python text mode): {univ_lines:,}")

    # Strict-newline line count: split only on \n
    strict_lf_lines = raw.count(b"\n") + (0 if raw.endswith(b"\n") else 1)
    print(f"Strict \\n line count (Git wc -l equivalent if no trailing \\n): {strict_lf_lines:,}")
    print(f"  (wc -l counts complete \\n-terminated lines, so it shows {n_lf:,})")

    # Trailing byte sanity
    last_bytes = raw[-20:]
    print(f"\nLast 20 bytes: {last_bytes!r}")

    # Unicode separators (decode-safe scan)
    text = raw.decode("utf-8", errors="replace")
    u2028 = text.count(" ")
    u2029 = text.count(" ")
    nel   = text.count("")
    print(f"\nUnicode line separators:")
    print(f"  U+2028 LINE SEPARATOR     : {u2028:,}")
    print(f"  U+2029 PARAGRAPH SEPARATOR: {u2029:,}")
    print(f"  U+0085 NEXT LINE          : {nel:,}")

    # NUL bytes (would break some readers, leave others alone)
    n_nul = raw.count(b"\x00")
    print(f"  NUL bytes                 : {n_nul:,}")

    # Power Query likely uses universal-newline + treats \r as line break too.
    # So expected PQ row count = univ_lines.
    print(f"\nReconciliation:")
    print(f"  wc -l                  reported : 246,522   (matches \\n count + EOF handling)")
    print(f"  Universal-newline      computed : {univ_lines:,}")
    print(f"  Power Query            reported : 250,715")
    print(f"  Gap (univ vs PQ)               : {250_715 - univ_lines:+,}")
    print(f"  Gap (wc-l vs PQ)               : {250_715 - 246_522:+,}")
    print(f"  Bare \\r count                  : {n_bare_cr:,}  <-- explains gap if non-zero")


if __name__ == "__main__":
    main()
