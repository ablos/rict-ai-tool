"""Write claims as JSONL (for machines) and as Excel review workbooks (for people).

The review is blind: reviewers first see only what the pipeline found (quotes with
their context), decide themselves what each quote says, and only then compare with
what the LLM decided. A review workbook covers one paper and has these sheets:

    Judge      a random sample of quotes; the reviewer picks a relation and writes the value
    Relations  the ontology's relations and definitions, for reference while judging
    Compare    the reviewer's decision next to the LLM's (open only after judging)
    Summary    agreement between reviewer and LLM, overall and per relation
    Run info   the settings of the run
"""

import json
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

NO_RELATION = "none"  # reviewer's choice when a quote says nothing the ontology covers
VALUE_MATCH = ["yes", "partly", "no"]

JUDGE_COLUMNS = ["item_id", "method", "quote", "context", "page", "section"]
JUDGE_REVIEW_COLUMNS = ["human_relation", "human_value", "notes", "reviewer"]
COMPARE_COLUMNS = ["item_id", "quote", "human_relation", "llm_relations", "relation_match",
                   "human_value", "llm_values", "value_match", "notes", "relation_key"]

# Column widths in characters; columns not listed get DEFAULT_WIDTH.
WIDTHS = {
    "method": 20, "quote": 60, "context": 80, "section": 26, "human_relation": 22, "human_value": 30,
    "notes": 30, "llm_relations": 22, "relation_match": 14, "llm_values": 30, "value_match": 12,
    "relation": 22, "concept": 22, "definition": 90,
}
DEFAULT_WIDTH = 12

GREY = PatternFill("solid", fgColor="E7E6E6")  # filled in automatically, not by the reviewer


def write_jsonl(rows: list[dict], path: str | Path) -> None:
    with open(path, "w", encoding="utf-8") as file:
        for row in rows:
            file.write(json.dumps(row, ensure_ascii=False) + "\n")


def read_jsonl(path: str | Path) -> list[dict]:
    with open(path, encoding="utf-8") as file:
        return [json.loads(line) for line in file if line.strip()]


def review_items(rows: list[dict]) -> list[dict]:
    """The claims grouped by quote: one item per (method, quote), with every LLM claim
    that quote supports. Reviewers judge quotes, so a quote the LLM used for three
    values is one item, not three."""
    items = {}

    for row in rows:
        if row["status"] != "reported":
            continue

        key = (row["method"], row["quote"])
        item = items.setdefault(key, {
            "method": row["method"], "quote": row["quote"], "context": row.get("context"),
            "page": row.get("page"), "section": row.get("section"), "claims": [],
        })
        item["claims"].append({"claim_id": row["claim_id"], "relation": row["relation"], "value": row["value"]})

    return list(items.values())


def sample_items(items: list[dict], size: int, seed: int) -> list[dict]:
    """A random sample spread over the relations: take items relation by relation in turn,
    so rare relations are checked too. size 0 means all items."""
    if size <= 0 or size >= len(items):
        return items

    rng = random.Random(seed)
    by_relation = {}

    for item in items:
        by_relation.setdefault(item["claims"][0]["relation"], []).append(item)

    for group in by_relation.values():
        rng.shuffle(group)

    sample = []

    while len(sample) < size:
        for group in by_relation.values():
            if group and len(sample) < size:
                sample.append(group.pop())

    rng.shuffle(sample)  # so reviewers don't see the relations in a pattern
    return sample


def write_review_sheet(rows: list[dict], path: str | Path, ontology: dict,
                       sample_size: int, seed: int, meta: dict | None = None) -> None:
    """A blind review workbook for the given rows (normally the rows of one paper)."""
    items = sample_items(review_items(rows), sample_size, seed)

    for number, item in enumerate(items, start=1):
        item["item_id"] = f"Q{number:02d}"

    relation_names = [relation["name"] for concept in ontology["concepts"] for relation in concept["relations"]]
    workbook = Workbook()

    judge = workbook.active
    judge.title = "Judge"
    _fill_sheet(judge, [[item.get(column) for column in JUDGE_COLUMNS] for item in items],
                JUDGE_COLUMNS + JUDGE_REVIEW_COLUMNS)

    reference = workbook.create_sheet("Relations")
    _fill_sheet(reference, [[relation["name"], concept["label"], relation["definition"]]
                            for concept in ontology["concepts"] for relation in concept["relations"]]
                + [[NO_RELATION, "", "The quote says nothing about the method that the ontology covers."]],
                ["relation", "concept", "definition"])
    # The dropdown reads its choices from the Relations sheet (a typed-out list is limited to 255 characters).
    _add_dropdown(judge, JUDGE_COLUMNS + JUDGE_REVIEW_COLUMNS, "human_relation",
                  f"Relations!$A$2:$A${len(relation_names) + 2}", len(items))

    compare = workbook.create_sheet("Compare")
    _add_compare(compare, items)

    summary = workbook.create_sheet("Summary")
    _add_summary(summary, relation_names, len(items))

    if meta:
        info = workbook.create_sheet("Run info")

        for key, value in _flatten(meta):
            info.append([key, value])

        info.column_dimensions["A"].width = 30
        info.column_dimensions["B"].width = 80

    workbook.save(path)


def _add_compare(sheet, items: list[dict]) -> None:
    """Human decision (copied from Judge by formula) next to the LLM's, row by row."""
    judge_headers = JUDGE_COLUMNS + JUDGE_REVIEW_COLUMNS
    human_relation = _letter(judge_headers, "human_relation")
    human_value = _letter(judge_headers, "human_value")
    data = []

    for row_number, item in enumerate(items, start=2):
        relations = "; ".join(dict.fromkeys(claim["relation"] for claim in item["claims"]))
        values = "\n".join(f"{claim['relation']}: {claim['value']}" for claim in item["claims"])
        # relation_key wraps every name in "; ...;", so "addresses" never matches inside
        # "addressesValue" when the formulas look a relation up.
        data.append([item["item_id"], item["quote"], f'=IF(Judge!{human_relation}{row_number}="","",'
                     f"Judge!{human_relation}{row_number})", relations, None,
                     f'=IF(Judge!{human_value}{row_number}="","",Judge!{human_value}{row_number})',
                     values, None, None, f"; {relations};"])

    _fill_sheet(sheet, data, COMPARE_COLUMNS)
    human, match, key = (_letter(COMPARE_COLUMNS, name) for name in ("human_relation", "relation_match", "relation_key"))

    for row_number in range(2, len(items) + 2):
        cell = sheet[f"{match}{row_number}"]
        cell.value = (f'=IF({human}{row_number}="","",IF(ISNUMBER(SEARCH("; "&{human}{row_number}&";",'
                      f'{key}{row_number})),"yes","no"))')

        for column in (human, match, _letter(COMPARE_COLUMNS, "human_value")):
            sheet[f"{column}{row_number}"].fill = GREY

    sheet.column_dimensions[key].hidden = True
    _add_dropdown(sheet, COMPARE_COLUMNS, "value_match", f'"{",".join(VALUE_MATCH)}"', len(items))


def _add_summary(sheet, relation_names: list[str], item_count: int) -> None:
    """Agreement between reviewer and LLM, with a 95% margin of error for the sample."""
    last = item_count + 1
    human, rel_match, val_match, llm = (f"Compare!${_letter(COMPARE_COLUMNS, name)}$2:${_letter(COMPARE_COLUMNS, name)}${last}"
                                        for name in ("human_relation", "relation_match", "value_match", "relation_key"))

    sheet.append(["Measure", "Value", "± (95%)"])
    sheet.append(["Quotes in the sample", item_count])
    sheet.append(["Quotes judged", f'=COUNTIF({rel_match},"yes")+COUNTIF({rel_match},"no")'])
    sheet.append(["Same relation as the LLM", f'=IFERROR(COUNTIF({rel_match},"yes")/B3,"")',
                  '=IFERROR(1.96*SQRT(B4*(1-B4)/B3),"")'])
    sheet.append(["Same relation and same value", f'=IFERROR(COUNTIFS({rel_match},"yes",{val_match},"yes")/B3,"")',
                  '=IFERROR(1.96*SQRT(B5*(1-B5)/B3),"")'])
    sheet.append(["Same relation, value partly the same",
                  f'=IFERROR(COUNTIFS({rel_match},"yes",{val_match},"partly")/B3,"")'])
    sheet.append(["Reviewer: quote supports nothing (LLM claim unsupported)",
                  f'=IFERROR(COUNTIF({human},"{NO_RELATION}")/B3,"")'])
    sheet.append([])
    sheet.append(["Relation (as chosen by the LLM)", "Quotes judged", "Same relation", "Same relation and value"])
    header_row = sheet.max_row

    for relation in relation_names:
        row = sheet.max_row + 1
        has_relation = f'"*; {relation};*"'
        sheet.append([
            relation,
            f'=COUNTIFS({llm},{has_relation},{rel_match},"yes")+COUNTIFS({llm},{has_relation},{rel_match},"no")',
            f'=IFERROR(COUNTIFS({llm},{has_relation},{rel_match},"yes")/B{row},"")',
            f'=IFERROR(COUNTIFS({llm},{has_relation},{rel_match},"yes",{val_match},"yes")/B{row},"")',
        ])

    for row_number in (1, header_row):
        for cell in sheet[row_number]:
            cell.font = Font(bold=True)

    for row_number in range(4, 8):
        for column in (2, 3):
            sheet.cell(row_number, column).number_format = "0%"

    for row_number in range(header_row + 1, sheet.max_row + 1):
        for column in (3, 4):
            sheet.cell(row_number, column).number_format = "0%"

    sheet.column_dimensions["A"].width = 52

    for column in "BCD":
        sheet.column_dimensions[column].width = 16


def _letter(headers: list[str], name: str) -> str:
    return get_column_letter(headers.index(name) + 1)


def _fill_sheet(sheet, data: list[list], headers: list[str]) -> None:
    """A header row and the data; columns without data stay empty for the reviewer."""
    sheet.append(headers)

    for values in data:
        sheet.append([_cell_value(value) for value in values] + [None] * (len(headers) - len(values)))

    for cell in sheet[1]:
        cell.font = Font(bold=True)

    for cells in sheet.iter_rows(min_row=2):
        for cell in cells:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    for index, name in enumerate(headers, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = WIDTHS.get(name, DEFAULT_WIDTH)

    sheet.freeze_panes = "B2"  # header row and first column stay visible while scrolling
    sheet.auto_filter.ref = sheet.dimensions


def _add_dropdown(sheet, headers: list[str], column: str, source: str, rows: int) -> None:
    """source is a quoted list ('"yes,no"') or a cell range ("Relations!$A$2:$A$9")."""
    letter = _letter(headers, column)
    dropdown = DataValidation(type="list", formula1=source, allow_blank=True)
    dropdown.add(f"{letter}2:{letter}{max(rows + 1, 2)}")
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
