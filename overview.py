"""Raw counts and all reviewer notes for one run, written as a Markdown overview.

    python overview.py runs/2026-10-17_173606     # writes runs/.../review_overview.md

Also written automatically by evaluate.py. Unlike evaluate.py, this gives no rates,
only the raw counts per paper and per relation, the screening decisions, and every
note reviewers wrote, for finding error patterns.
"""

import sys
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook

from export import JUDGE_COLUMNS, JUDGE_REVIEW_COLUMNS, SUPPORTS_NOTHING, relation_from_label

PAPER_COLUMNS = [
    ("quotes", "Quotes"), ("supports_nothing", "Supports nothing"), ("exact", "Exact match"),
    ("llm_rel", "LLM relations"), ("human_rel", "Human relations"), ("agreed", "Agreed"),
    ("llm_only", "LLM only"), ("llm_only_correct", "LLM only, right after all"),
    ("human_only", "Human only"), ("value_yes", "Value yes"), ("value_partly", "Value partly"),
    ("value_no", "Value no"), ("value_empty", "Value empty"),
]


def write_overview(run_dir: Path, out_path: Path) -> None:
    """Write the raw counts, screening table and all notes of a run to out_path (Markdown)."""
    paper_rows, per_relation, notes = _score_reviews(run_dir)
    totals = sum((counts for _, counts in paper_rows), Counter())
    out = [f"# Review overview: {run_dir.name}\n", "## Raw counts per paper\n",
           "| Paper | " + " | ".join(label for _, label in PAPER_COLUMNS) + " |",
           "|---|" + "---:|" * len(PAPER_COLUMNS)]

    for paper, counts in paper_rows + [("**Total**", totals)]:
        out.append(f"| {paper} | " + " | ".join(str(counts.get(key, 0)) for key, _ in PAPER_COLUMNS) + " |")

    out.append("\nLLM relations = agreed + LLM only; human relations = agreed + human only. Counts are "
               "(quote, relation) pairs, except Quotes, Supports nothing, Exact match and the Value columns (quotes).\n")
    out += ["## Raw counts per relation (all papers)\n",
            "| Relation | LLM | Human | Agreed | LLM only | Human only |", "|---|---:|---:|---:|---:|---:|"]

    for name, rel in sorted(per_relation.items(), key=lambda item: -(item[1]["llm_only"] + item[1]["human_only"])):
        out.append(f"| {name} | {rel['llm']} | {rel['human']} | {rel['agreed']} | {rel['llm_only']} | {rel['human_only']} |")

    out += _screening_section(run_dir / "screening.xlsx")
    out.append(f"\n## All notes ({len(notes)})\n")
    current = None

    for paper, item, where, quote, human, llm, note in notes:
        if paper != current:
            out.append(f"\n### {paper}\n")
            current = paper

        out.append(f"- **{item}** ({where}) — human: {human}; LLM: {llm}\n  - Quote: \"{quote}\"\n  - Note: {note}")

    out_path.write_text("\n".join(out) + "\n", encoding="utf-8")


def _score_reviews(run_dir: Path):
    """Counts per paper (fully or partly judged papers only), per relation, and all notes."""
    paper_rows, per_relation, notes = [], {}, []

    for path in sorted(run_dir.glob("review_*.xlsx")):
        workbook = load_workbook(path)
        judge, compare_sheet = workbook["Judge"], workbook["Compare"]
        judge_headers = [cell.value for cell in judge[1]]
        compare_headers = [cell.value for cell in compare_sheet[1]]
        labels = judge_headers[len(JUDGE_COLUMNS):len(judge_headers) - len(JUDGE_REVIEW_COLUMNS)]
        compare = {row[0]: dict(zip(compare_headers, row))
                   for row in compare_sheet.iter_rows(min_row=2, values_only=True) if row[0]}
        paper = path.name.removeprefix("review_").removesuffix(".xlsx")
        counts = Counter()

        for values in judge.iter_rows(min_row=2, values_only=True):
            row = dict(zip(judge_headers, values))

            if not row.get("item_id"):
                continue

            human = {relation_from_label(label) for label in labels if _text(row.get(label))}

            if not human and not any(_text(row.get(name)) for name in (SUPPORTS_NOTHING, "human_value", "reviewer")):
                counts["not_judged"] += 1
                continue

            compared = compare.get(row["item_id"], {})
            llm = {name.strip() for name in _text(compared.get("llm_relations")).split(";") if name.strip()}
            agreed, llm_only, human_only = human & llm, llm - human, human - llm
            correct = _count(compared.get("llm_only_correct"))
            counts.update(quotes=1, supports_nothing=int(not human), exact=int(human == llm),
                          llm_rel=len(llm), human_rel=len(human), agreed=len(agreed),
                          llm_only=len(llm_only), human_only=len(human_only),
                          llm_only_correct=min(correct, len(llm_only)))

            if agreed:
                counts[f"value_{_text(compared.get('value_ok')).lower() or 'empty'}"] += 1

            for name in llm | human:
                per_relation.setdefault(name, Counter()).update(
                    llm=int(name in llm), human=int(name in human), agreed=int(name in agreed),
                    llm_only=int(name in llm_only), human_only=int(name in human_only))

            for where, note in (("Judge", row.get("notes")), ("Compare", compared.get("notes"))):
                if _text(note):
                    notes.append((paper, row["item_id"], where, _text(row.get("quote"))[:160],
                                  ", ".join(sorted(human)) or "(nothing)", ", ".join(sorted(llm)) or "(nothing)",
                                  _text(note)))

        if counts["quotes"]:
            paper_rows.append((paper, counts))

    return paper_rows, per_relation, notes


def _screening_section(path: Path) -> list[str]:
    """The screening decisions per method, with disagreements marked."""
    if not path.exists():
        return []

    workbook = load_workbook(path)
    judge_headers = [cell.value for cell in workbook["Judge"][1]]
    compare_headers = [cell.value for cell in workbook["Compare"][1]]
    llm = {(row[0], row[1]): dict(zip(compare_headers, row))
           for row in workbook["Compare"].iter_rows(min_row=2, values_only=True) if row[0]}
    summary, rows = Counter(), []

    for values in workbook["Judge"].iter_rows(min_row=2, values_only=True):
        row = dict(zip(judge_headers, values))

        if not row.get("paper"):
            continue

        human = _text(row.get("human_decision")).lower()
        compared = llm.get((row["paper"], row["method"]))
        llm_decision = _text(compared.get("llm_decision")).lower() if compared else "(missed by LLM)"

        if human:
            summary["judged"] += 1
            summary[f"LLM {llm_decision} / reviewer {human}"] += 1
        else:
            summary["not judged"] += 1

        mark = " **≠**" if human and human != llm_decision else ""
        note = _text(row.get("notes")) or _text((compared or {}).get("notes"))
        # Per criterion: reviewer / LLM, e.g. "met / not met".
        criteria = " · ".join(f"{c}: {_text(row.get(c)) or '-'} / {_text((compared or {}).get(f'llm_{c}')) or '-'}"
                              for c in ("C1", "C2", "C3"))
        rows.append(f"| {row['paper']} | {row['method']} | {llm_decision} | {human or '-'}{mark} | {criteria} | {note} |")

    return (["\n## Screening\n"] + [f"- {key}: {value}" for key, value in sorted(summary.items())]
            + ["\n| Paper | Method | LLM | Reviewer | Criteria (reviewer / LLM) | Notes |", "|---|---|---|---|---|---|"]
            + rows)


def _text(value) -> str:
    return " ".join(str(value or "").split())


def _count(value) -> int:
    """A whole number typed by a reviewer; empty or unreadable counts as 0."""
    try:
        return max(0, int(float(value)))
    except (TypeError, ValueError):
        return 0


if __name__ == "__main__":
    run = Path(sys.argv[1])
    write_overview(run, run / "review_overview.md")
    print(f"Written to {run / 'review_overview.md'}")
