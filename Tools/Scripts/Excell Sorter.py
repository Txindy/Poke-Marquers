#!/usr/bin/env python3
import pandas as pd
import argparse
from pathlib import Path
import sys
from typing import Optional

def main(path_str: Optional[str]) -> None:
    # Resolve base directory for script vs frozen exe
    if getattr(sys, 'frozen', False):
        base_dir = Path(sys.executable).resolve().parent
    else:
        base_dir = Path(__file__).resolve().parent

    # Determine target Excel file
    if path_str:
        path = Path(path_str)
        if not path.is_absolute():
            path = base_dir / path
    else:
        # Find any Excel file in the same folder (skip temporary '~$' files)
        candidates = [
            p for p in base_dir.iterdir()
            if p.is_file() and not p.name.startswith('~$') and p.suffix.lower() in {'.xlsx', '.xlsm'}
        ]
        if not candidates:
            print("Error: no Excel files (.xlsx, .xlsm) found next to the script/exe.")
            sys.exit(1)
        # Pick the most recently modified file
        path = max(candidates, key=lambda p: p.stat().st_mtime)
        print(f"ℹ️ Using detected file: {path.name}")

    if not path.exists():
        print(f"Error: file not found: {path}")
        sys.exit(1)

    df = pd.read_excel(path)

    required = {"Variant Type", "Number"}
    missing = required - set(df.columns)
    if missing:
        print(f"Error: missing columns: {', '.join(sorted(missing))}")
        sys.exit(1)

    variant_order = {"Normal": 0, "Reverse Holo": 1}
    df["VariantSort"] = df["Variant Type"].map(variant_order)

    df.sort_values(by=["Number", "VariantSort"], inplace=True)
    df.drop(columns=["VariantSort"], inplace=True)

    df.to_excel(path, index=False)
    print("✅ File has been reordered and saved!")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Reorder Excel by Number and Variant Type.")
    parser.add_argument("file", nargs="?", default=None,
                        help="Path to the Excel file. If omitted, the script searches for any .xlsx/.xlsm next to it.")
    args = parser.parse_args()
    main(args.file)