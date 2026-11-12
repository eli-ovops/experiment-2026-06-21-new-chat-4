from __future__ import annotations

import csv
import json
import re
from datetime import date, datetime, time
from pathlib import Path

from openpyxl import load_workbook


SOURCE = Path(
    "/Users/USER/Library/Containers/com.tencent.xinWeChat/Data/Documents/"
    "xwechat_files/wxid_ahypc7nu2jlp22_98f0/temp/drag/"
    "5S客人6月2日实装（May23,26).xlsx"
)
OUTPUT_DIR = Path("/Users/USER/Documents/Codex/2026-06-21/new-chat-4/outputs")


def safe_name(value: str) -> str:
    value = re.sub(r'[\\/:*?"<>|]+', "_", value).strip()
    return value or "Sheet"


def csv_value(data_value, formula_value):
    value = data_value
    if value is None and isinstance(formula_value, str) and formula_value.startswith("="):
        value = formula_value
    if value is None:
        return ""
    if isinstance(value, datetime):
        if value.time() == time(0, 0):
            return value.date().isoformat()
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    return value


def trim_trailing_empty(rows):
    trimmed_rows = []
    max_cols = 0
    for row in rows:
        row_values = list(row)
        while row_values and row_values[-1] == "":
            row_values.pop()
        trimmed_rows.append(row_values)
        max_cols = max(max_cols, len(row_values))

    while trimmed_rows and not trimmed_rows[-1]:
        trimmed_rows.pop()

    return [row + [""] * (max_cols - len(row)) for row in trimmed_rows], max_cols


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    wb_values = load_workbook(SOURCE, read_only=True, data_only=True)
    wb_formulas = load_workbook(SOURCE, read_only=True, data_only=False)

    base_name = SOURCE.stem
    multi_sheet = len(wb_values.sheetnames) > 1
    results = []

    for sheet_name in wb_values.sheetnames:
        ws_values = wb_values[sheet_name]
        ws_formulas = wb_formulas[sheet_name]
        rows = []
        for data_row, formula_row in zip(ws_values.iter_rows(), ws_formulas.iter_rows()):
            rows.append(
                [
                    csv_value(data_cell.value, formula_cell.value)
                    for data_cell, formula_cell in zip(data_row, formula_row)
                ]
            )

        rows, cols = trim_trailing_empty(rows)
        if multi_sheet:
            output_path = OUTPUT_DIR / f"{base_name} - {safe_name(sheet_name)}.csv"
        else:
            output_path = OUTPUT_DIR / f"{base_name}.csv"

        with output_path.open("w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.writer(handle)
            writer.writerows(rows)

        results.append(
            {
                "sheet": sheet_name,
                "csv": str(output_path),
                "rows": len(rows),
                "columns": cols,
            }
        )

    manifest_path = OUTPUT_DIR / f"{base_name} - csv_manifest.json"
    manifest_path.write_text(
        json.dumps(
            {
                "source": str(SOURCE),
                "sheet_count": len(wb_values.sheetnames),
                "outputs": results,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(json.dumps({"manifest": str(manifest_path), "outputs": results}, ensure_ascii=False))


if __name__ == "__main__":
    main()
