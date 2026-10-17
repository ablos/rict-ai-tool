"""Add up the reviewed workbooks of one or more runs into an accuracy estimate.

    python evaluate.py runs/2026-10-17_101500                         # one run
    python evaluate.py runs/2026-10-17_101500 runs/2026-10-20_143000  # compare two runs

Reads every review_*.xlsx (and screening.xlsx, if present) in each run folder,
compares the reviewers' decisions (Judge sheet) with the LLM's (Compare sheet), and
prints how often they agree, overall and per relation, with a 95% confidence
interval, plus the include/exclude agreement of the screening. Saves the result as
scores.json in the run folder. Everything is recomputed from what the reviewers
typed, so it does not depend on the spreadsheet having calculated its formulas.
"""

import json
import math
import sys
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import load_workbook

from export import NO_RELATION, read_jsonl, review_items


def score_run(run_dir: Path) -> dict:
    """Agreement between reviewers and LLM for all review workbooks in a run folder."""
    totals, per_relation, per_paper, warnings = Counter(), defaultdict(Counter), {}, []
    # How many quotes each paper has in total (the pool its sample was drawn from), for weighting.
    rows = read_jsonl(run_dir / "claims.jsonl")
    quotes_per_paper = {
        f"review_{Path(publication).stem}.xlsx": len(review_items([row for row in rows if row["publication"] == publication]))
        for publication in dict.fromkeys(row["publication"] for row in rows)
    }

    for path in sorted(run_dir.glob("review_*.xlsx")):
        workbook = load_workbook(path)
        judged = list(_rows(workbook["Judge"]))
        compared = {row["item_id"]: row for row in _rows(workbook["Compare"])}
        paper = Counter()

        for row in judged:
            human_relation = (row["human_relation"] or "").strip()

            if not human_relation:
                paper["not_judged"] += 1
                continue

            llm_relations = [name.strip() for name in (compared[row["item_id"]]["llm_relations"] or "").split(";")]
            value_match = (compared[row["item_id"]]["value_match"] or "").strip()
            outcome = Counter(judged=1)
            outcome["same_relation"] = human_relation in llm_relations
            outcome["same_relation_and_value"] = outcome["same_relation"] and value_match == "yes"
            outcome["value_partly"] = outcome["same_relation"] and value_match == "partly"
            outcome["value_not_compared"] = outcome["same_relation"] and not value_match
            outcome["human_none"] = human_relation == NO_RELATION
            paper += outcome

            for relation in llm_relations:
                per_relation[relation] += outcome

        if paper["not_judged"] or paper["value_not_compared"]:
            warnings.append(f"{path.name}: {paper['not_judged']} quotes not judged, "
                            f"{paper['value_not_compared']} values not compared yet")

        per_paper[path.name] = _with_rates(paper)
        per_paper[path.name]["quotes_in_paper"] = quotes_per_paper.get(path.name, 0)
        totals += paper

    total = _with_rates(totals)
    total["weighted"] = {key: _weighted(per_paper, key) for key in ("same_relation", "same_relation_and_value")}

    return {
        "run": str(run_dir),
        "total": total,
        "per_paper": per_paper,
        "per_relation": {relation: _with_rates(counts) for relation, counts in per_relation.items()},
        "screening": score_screening(run_dir / "screening.xlsx", warnings),
        "warnings": warnings,
    }


def score_screening(path: Path, warnings: list[str]) -> dict | None:
    """Include/exclude agreement from screening.xlsx (if the run has one), per method and per paper.

    Rows the reviewer added on the Judge sheet are methods the LLM did not find:
    marked "include", they count as missed by the LLM.
    """
    if not path.exists():
        return None

    workbook = load_workbook(path)
    llm = {(row["paper"], row["method"]): row["llm_decision"] for row in _rows(workbook["Compare"])}
    methods = Counter()
    papers = defaultdict(lambda: {"human": False, "llm": False, "judged": False})

    for (paper, _), decision in llm.items():
        papers[paper]["llm"] |= decision == "include"

    for row in _rows(workbook["Judge"]):
        human = (row["human_decision"] or "").strip()
        key = (row["paper"], row["method"])

        if not human:
            methods["not_judged"] += 1
            continue

        papers[row["paper"]]["judged"] = True
        papers[row["paper"]]["human"] |= human == "include"

        if key not in llm:  # added by the reviewer
            methods["missed_by_llm"] += human == "include"
            continue

        methods["judged"] += 1
        methods["same_decision"] += human == llm[key]
        methods[f"llm_{llm[key]}_human_{human}"] += 1

    if methods["not_judged"]:
        warnings.append(f"screening.xlsx: {methods['not_judged']} methods not judged")

    paper_counts = Counter()

    for paper in papers.values():
        if paper["judged"]:
            paper_counts["judged"] += 1
            paper_counts["same_decision"] += paper["human"] == paper["llm"]
            paper_counts[f"llm_{'include' if paper['llm'] else 'exclude'}_human_{'include' if paper['human'] else 'exclude'}"] += 1

    methods = dict(methods)
    methods["same_decision_rate"] = _wilson(methods.get("same_decision", 0), methods.get("judged", 0))
    paper_counts = dict(paper_counts)
    paper_counts["same_decision_rate"] = _wilson(paper_counts.get("same_decision", 0), paper_counts.get("judged", 0))
    return {"per_method": methods, "per_paper": paper_counts}


def _with_rates(counts: Counter) -> dict:
    """Counts plus agreement rates with a 95% Wilson confidence interval (sound for small samples)."""
    result = dict(counts)
    n = counts["judged"]

    for key in ("same_relation", "same_relation_and_value"):
        result[f"{key}_rate"] = _wilson(counts[key], n)

    return result


def _weighted(per_paper: dict, key: str, z: float = 1.96) -> dict | None:
    """Agreement rate per quote in the whole run: each paper counts in proportion to its
    number of quotes, not equally (stratified estimate, one stratum per paper).

    The margin uses the stratified-sampling variance with a finite population correction,
    since a sample of 20 can be a large share of a small paper.
    """
    papers = [paper for paper in per_paper.values() if paper.get("judged") and paper.get("quotes_in_paper")]
    total_quotes = sum(paper["quotes_in_paper"] for paper in papers)

    if not total_quotes:
        return None

    rate, variance = 0.0, 0.0

    for paper in papers:
        weight = paper["quotes_in_paper"] / total_quotes
        n, size = paper["judged"], paper["quotes_in_paper"]
        p = paper.get(key, 0) / n
        rate += weight * p
        correction = (size - n) / (size - 1) if size > 1 else 0.0
        variance += weight ** 2 * p * (1 - p) / n * correction

    margin = z * math.sqrt(variance)
    return {"rate": round(rate, 3), "low": round(max(0.0, rate - margin), 3), "high": round(min(1.0, rate + margin), 3)}


def _wilson(successes: int, n: int, z: float = 1.96) -> dict | None:
    if n == 0:
        return None

    p = successes / n
    centre = (p + z * z / (2 * n)) / (1 + z * z / n)
    margin = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / (1 + z * z / n)
    return {"rate": round(p, 3), "low": round(centre - margin, 3), "high": round(centre + margin, 3)}


def _rows(sheet):
    """The sheet's rows as dicts keyed by the header row."""
    rows = sheet.iter_rows(values_only=True)
    headers = next(rows)

    for values in rows:
        if any(value is not None for value in values):
            yield dict(zip(headers, values))


def _rate(rate: dict | None) -> str:
    return "-" if rate is None else f"{rate['rate']:.0%} ({rate['low']:.0%}-{rate['high']:.0%})"


def print_scores(scores: list[dict]) -> None:
    width, column = 40, 24
    print(f"{'':<{width}}" + "".join(f"{Path(s['run']).name:>{column}}" for s in scores))

    for key, label in [("judged", "Quotes judged"), ("not_judged", "Quotes not judged yet"),
                       ("human_none", "Reviewer: quote supports nothing")]:
        print(f"{label:<{width}}" + "".join(f"{s['total'].get(key, 0):>{column}}" for s in scores))

    print("\nEvery paper counts equally:")

    for key, label in [("same_relation_rate", "Same relation (95% CI)"),
                       ("same_relation_and_value_rate", "Same relation and value (95% CI)")]:
        print(f"  {label:<{width - 2}}" + "".join(f"{_rate(s['total'].get(key)):>{column}}" for s in scores))

    print("Weighted by each paper's number of quotes (accuracy per quote in the knowledge base):")

    for key, label in [("same_relation", "Same relation (95% CI)"),
                       ("same_relation_and_value", "Same relation and value (95% CI)")]:
        print(f"  {label:<{width - 2}}" + "".join(f"{_rate(s['total']['weighted'].get(key)):>{column}}" for s in scores))

    print(f"\nPer relation chosen by the LLM: same relation and value (quotes judged)")
    relations = dict.fromkeys(relation for s in scores for relation in s["per_relation"])

    for relation in relations:
        cells = []

        for s in scores:
            counts = s["per_relation"].get(relation)
            cells.append("-" if not counts else
                         f"{_rate(counts['same_relation_and_value_rate'])} ({counts['judged']})")

        print(f"  {relation:<{width - 2}}" + "".join(f"{cell:>{column}}" for cell in cells))

    if any(s["screening"] for s in scores):
        for level, title, extra in [("per_method", "Screening per method", [("missed_by_llm", "Missed by the LLM (added by reviewer)")]),
                                    ("per_paper", "Screening per paper (included if at least one method is)", [])]:
            print(f"\n{title}")

            for key, label in [("judged", "Judged"),
                               ("llm_include_human_exclude", "LLM included, reviewer excluded"),
                               ("llm_exclude_human_include", "LLM excluded, reviewer included")] + extra:
                print(f"  {label:<{width - 2}}" + "".join(
                    f"{s['screening'][level].get(key, 0) if s['screening'] else '-':>{column}}" for s in scores))

            print(f"  {'Same decision (95% CI)':<{width - 2}}" + "".join(
                f"{_rate(s['screening'][level]['same_decision_rate']) if s['screening'] else '-':>{column}}" for s in scores))

    for s in scores:
        for warning in s["warnings"]:
            print(f"\nWarning ({Path(s['run']).name}): {warning}")


def main() -> None:
    run_dirs = [Path(arg) for arg in sys.argv[1:]]

    if not run_dirs:
        sys.exit(__doc__)

    scores = []

    for run_dir in run_dirs:
        if not list(run_dir.glob("review_*.xlsx")) and not (run_dir / "screening.xlsx").exists():
            sys.exit(f"No review_*.xlsx or screening.xlsx in {run_dir}")

        result = score_run(run_dir)
        (run_dir / "scores.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
        scores.append(result)

    print_scores(scores)


if __name__ == "__main__":
    main()
