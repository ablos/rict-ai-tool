"""Write claims as JSONL (for machines) and as an Excel review sheet (for people)."""

import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

CLAIM_VERDICTS = ["correct", "incorrect", "hallucinated", "vague", "unsupported inference"]
NOT_REPORTED_VERDICTS = ["correct (not in paper)", "missed (in paper)"]
REVIEW_COLUMNS = ["verdict", "correction", "notes", "reviewer"]

CLAIM_COLUMNS = [
    "claim_id", "publication", "method", "relation", "concept", "value",
    "quote", "page", "section", "context", "quote_found", "match_score", "found_page", "flags",
]
NOT_REPORTED_COLUMNS = ["claim_id", "publication", "method", "relation", "concept"]

# Column widths in characters; columns not listed get DEFAULT_WIDTH.
WIDTHS = {
    "publication": 24, "method": 20, "relation": 20, "concept": 18, "value": 30,
    "quote": 60, "section": 26, "context": 80, "flags": 30, "verdict": 22, "correction": 30, "notes": 30,
}
DEFAULT_WIDTH = 12

RED = PatternFill("solid", fgColor="F8CBAD")  # quote not found in the paper: likely hallucinated
YELLOW = PatternFill("solid", fgColor="FFE699")  # other flags: check with extra care


def write_jsonl(rows: list[dict], path: str | Path) -> None:
    with open(path, "w", encoding="utf-8") as file:
        for row in rows:
            file.write(json.dumps(row, ensure_ascii=False) + "\n")


def read_jsonl(path: str | Path) -> list[dict]:
    with open(path, encoding="utf-8") as file:
        return [json.loads(line) for line in file if line.strip()]


def write_review_sheet(rows: list[dict], path: str | Path, meta: dict | None = None) -> None:
    """An Excel file with a sheet of claims, a sheet of not-reported relations, and the run info."""
    claims = [row for row in rows if row["status"] == "reported"]
    # Quotes not found first, then other flagged claims, then the rest.
    # sort() keeps the original order within each group.
    claims.sort(key=lambda row: (row.get("quote_found") is not False, not row.get("flags")))
    not_reported = [row for row in rows if row["status"] == "not_reported"]

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Claims"
    _fill_sheet(sheet, claims, CLAIM_COLUMNS, CLAIM_VERDICTS)

    for row_number, claim in enumerate(claims, start=2):
        fill = RED if claim.get("quote_found") is False else YELLOW if claim.get("flags") else None
        
        if fill:
            for cell in sheet[row_number]:
                cell.fill = fill

    _fill_sheet(workbook.create_sheet("Not reported"), not_reported, NOT_REPORTED_COLUMNS, NOT_REPORTED_VERDICTS)

    if meta:
        info = workbook.create_sheet("Run info")
        
        for key, value in _flatten(meta):
            info.append([key, value])

        info.column_dimensions["A"].width = 30
        info.column_dimensions["B"].width = 80

    workbook.save(path)


def _fill_sheet(sheet, rows: list[dict], columns: list[str], verdicts: list[str]) -> None:
    """Data columns, then empty review columns with a dropdown for the verdict."""
    headers = columns + REVIEW_COLUMNS
    sheet.append(headers)
    
    for row in rows:
        sheet.append([_cell_value(row.get(column)) for column in columns] + [None] * len(REVIEW_COLUMNS))

    for cell in sheet[1]:
        cell.font = Font(bold=True)

    for cells in sheet.iter_rows(min_row=2):
        for cell in cells:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    for index, name in enumerate(headers, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = WIDTHS.get(name, DEFAULT_WIDTH)

    sheet.freeze_panes = "B2"  # header row and claim_id stay visible while scrolling
    sheet.auto_filter.ref = sheet.dimensions

    verdict_column = get_column_letter(headers.index("verdict") + 1)
    dropdown = DataValidation(type="list", formula1=f'"{",".join(verdicts)}"', allow_blank=True)
    dropdown.add(f"{verdict_column}2:{verdict_column}{max(len(rows) + 1, 2)}")
    sheet.add_data_validation(dropdown)


def _cell_value(value):
    """Excel cells hold text, numbers and booleans; anything else is written as JSON."""
    if value is None or isinstance(value, (str, int, float, bool)):
        return value

    return json.dumps(value, ensure_ascii=False)


def _flatten(data: dict, prefix: str = ""):
    """Yield ("a.b", value) pairs from a nested dict, for the Run info sheet."""
    for key, value in data.items():
        if isinstance(value, dict):
            yield from _flatten(value, f"{prefix}{key}.")
            
        else:
            yield f"{prefix}{key}", _cell_value(value)
