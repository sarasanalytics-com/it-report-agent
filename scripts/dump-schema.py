#!/usr/bin/env python3
"""Print the schema (sheet names + column headers) of the downloaded IT
workbooks, so the column mapping can be verified against the real files.

Privacy: prints only the detected HEADER row of each sheet (column names) and a
row count — never data rows — so no employee/PII values are logged.

Run after scripts/fetch-excel.py has populated data/.
"""

import pathlib

import openpyxl

DATA_DIR = pathlib.Path(__file__).resolve().parent.parent / "data"

FILES = [
    "asset_inventory.xlsx",
    "spend_tracker.xlsx",
    "procurement_plan.xlsx",
    "joiners_info.xlsx",
    "vendor_payments.xlsx",
]


def _header_row(ws):
    """Pick the most header-like row in the first 6 rows (most non-empty cells)."""
    best_idx, best_cells, best_count = None, [], -1
    for i, raw in enumerate(ws.iter_rows(min_row=1, max_row=6, values_only=True), start=1):
        cells = [str(c).strip() for c in raw if c not in (None, "")]
        if len(cells) > best_count:
            best_idx, best_cells, best_count = i, cells, len(cells)
    return best_idx, best_cells


def main() -> None:
    for fname in FILES:
        path = DATA_DIR / fname
        print(f"\n{'=' * 70}\n# {fname}")
        if not path.exists():
            print("  (not downloaded)")
            continue
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        except Exception as exc:  # noqa: BLE001
            print(f"  ERROR opening: {exc}")
            continue
        print(f"  sheets: {wb.sheetnames}")
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            hdr_idx, headers = _header_row(ws)
            n_rows = sum(1 for _ in ws.iter_rows(min_row=1, values_only=True))
            print(f"\n  -- sheet: {sheet!r}  (~{n_rows} rows, header looks like row {hdr_idx})")
            print(f"     columns: {headers}")
        wb.close()


if __name__ == "__main__":
    main()
