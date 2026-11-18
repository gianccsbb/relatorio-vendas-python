"""Command line entry point.

    python -m sales_report data/raw                      # every CSV/XLSX in the folder
    python -m sales_report data/raw -o reports/out.xlsx  # choose the output file
"""
import argparse
import sys
from pathlib import Path

from .cleaning import clean
from .report import build_report

SUPPORTED = {".csv", ".xlsx", ".xls"}


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="sales_report", description="Clean messy sales exports and build an Excel report.")
    parser.add_argument("inputs", nargs="+", help="Files or folders with CSV/XLSX exports")
    parser.add_argument("-o", "--output", default="reports/relatorio-vendas.xlsx", help="Output .xlsx path")
    parser.add_argument("--csv", help="Also save the clean data as CSV to this path")
    args = parser.parse_args(argv)

    files = []
    for item in map(Path, args.inputs):
        if item.is_dir():
            files += sorted(p for p in item.iterdir() if p.suffix.lower() in SUPPORTED and not p.name.startswith("~$"))
        elif item.suffix.lower() in SUPPORTED:
            files.append(item)
    if not files:
        print("No CSV/XLSX files found.", file=sys.stderr)
        return 1

    print(f"Reading {len(files)} file(s): " + ", ".join(f.name for f in files))
    df, log = clean(files)
    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    stats = build_report(df, log.to_frame(), str(out))
    if args.csv:
        df.to_csv(args.csv, index=False, encoding="utf-8-sig")
    print(f"{stats['rows']} clean sales · {stats['products']} products · {stats['days']} days · "
          f"{stats['corrections']} corrections logged")
    print(f"Report saved to {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
