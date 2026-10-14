#!/usr/bin/env python3
"""Check a dedup.py report against a list of known pair outcomes.

The expectations CSV holds real DUIDs, so it lives in the data directory,
never in this repository.  Columns:

  duid_a,duid_b,expect,note

where expect is "together" (same duplicate group) or "apart" (never in the
same group), and note is a short description with no names.

Usage:
  tools/check_pairs.py REPORT.xlsx EXPECTATIONS.csv

Exit status is 1 if any expectation fails.
"""

import argparse
import csv
import sys

import openpyxl


def groups(path):
    """Map DUID -> group number from the report's Duplicates sheet."""
    wb = openpyxl.load_workbook(path, read_only=True)
    rows = wb["Duplicates"].iter_rows(values_only=True)
    hdr = {h: i for i, h in enumerate(next(rows))}
    return {r[hdr["DUID"]]: r[0] for r in rows}


def main():
    """Compare each expectation with the report and print a pass/fail line."""
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("report")
    ap.add_argument("expectations")
    args = ap.parse_args()

    grp = groups(args.report)
    failures = 0
    with open(args.expectations, newline="") as fh:
        for row in csv.DictReader(fh):
            a, b = int(row["duid_a"]), int(row["duid_b"])
            together = grp.get(a) is not None and grp.get(a) == grp.get(b)
            ok = together if row["expect"] == "together" else not together
            failures += not ok
            print(f"{'ok  ' if ok else 'FAIL'} {a} {b} expect {row['expect']:<8} {row.get('note', '')}")
    print(f"\n{failures} failure(s)")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
