"""Check barcode lengths and auto-detect the proper (most common) length.

Usage:
    python length_check.py "path/to/report.xlsx"
    python length_check.py "path/to/report.xlsx" "QR Code"   # optional column name

It prints the length distribution, the detected proper length, and saves the
rows whose length differs to "<file>_length_offenders.xlsx".
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from barcode_tools import io as btio  # noqa: E402


def analyze(path, column=None):
    df = btio.read_table(path)
    if df.empty:
        print("No rows found.")
        return None

    if column is None:
        column = btio.guess_code_column(df)
    if column not in df.columns:
        print(f"Column '{column}' not found. Available: {list(df.columns)}")
        return None

    values = df[column]
    lengths = btio.code_lengths(values)
    proper = btio.guess_code_length(values)
    outliers = lengths != proper
    outlier_count = int(outliers.sum())

    print(f"File:                   {path}")
    print(f"Column:                 {column}")
    print(f"Rows:                   {len(df):,}")
    print(f"Detected proper length: {proper}")
    print(f"Outliers:               {outlier_count:,}")
    print()
    print("Length distribution (length -> count):")
    print(btio.length_distribution(values).to_string(index=False))
    print()

    if outlier_count:
        offenders = df[outliers].copy()
        offenders.insert(0, "Excel Row", offenders.index + 2)
        offenders["_Length"] = lengths[outliers].values
        out_path = os.path.splitext(path)[0] + "_length_offenders.xlsx"
        btio.write_excel(offenders, out_path)
        print(f"Saved {outlier_count:,} offenders to: {out_path}")
    else:
        print("All rows match the detected length.")

    return proper


def main():
    if len(sys.argv) > 1:
        path = sys.argv[1]
    else:
        path = input("Path to report file: ").strip().strip('"')
    column = sys.argv[2] if len(sys.argv) > 2 else None
    analyze(path, column)


if __name__ == "__main__":
    main()
