"""Add up the reviewed workbooks of one or more runs into an accuracy estimate.

    python evaluate.py runs/2026-10-17_101500                         # one run
    python evaluate.py runs/2026-10-17_101500 runs/2026-10-20_143000  # compare two runs

Reads every review_*.xlsx (and screening.xlsx, if present) in each run folder,
compares the reviewers' relations per quote (Judge sheet) with the LLM's (Compare
sheet), and prints exact agreement, relation precision and recall, and how often the
LLM's values are right, overall and per relation, plus the include/exclude agreement
of the screening. Saves the result as
scores.json in the run folder. Everything is recomputed from what the reviewers
typed, so it does not depend on the spreadsheet having calculated its formulas.
"""

import json
import math
import sys
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import load_workbook

from export import JUDGE_COLUMNS, JUDGE_REVIEW_COLUMNS, SUPPORTS_NOTHING, read_jsonl, relation_from_label, review_items


def score_run(run_dir: Path) -> dict:
    """Agreement between reviewers and LLM for all review workbooks in a run folder.

    Per quote, the reviewer's set of relations is compared with the LLM's:
        exact       both sets are the same
        precision   share of the LLM's (quote, relation) pairs the reviewer also gave
        recall      share of the reviewer's pairs the LLM also gave
        values      share of quotes with an agreed relation whose LLM value is right
    """
    totals, per_relation, per_paper, warnings = Counter(), defaultdict(Counter), {}, []
    # How many quotes each paper has in total (the pool its sample was drawn from), for weighting.
    rows = read_jsonl(run_dir / "claims.jsonl")
    quotes_per_paper = {
        f"review_{Path(publication).stem}.xlsx": len(review_items([row for row in rows if row["publication"] == publication]))
        for publication in dict.fromkeys(row["publication"] for row in rows)
    }

    for path in sorted(run_dir.glob("review_*.xlsx")):
        workbook = load_workbook(path)
        headers = [cell.value for cell in workbook["Judge"][1]]
        # The relation columns are labelled "relation (concept)".
        relation_labels = headers[len(JUDGE_COLUMNS):len(headers) - len(JUDGE_REVIEW_COLUMNS)]
        compared = {row["item_id"]: row for row in _rows(workbook["Compare"])}
        paper = Counter()

        for row in _rows(workbook["Judge"]):
            human = {relation_from_label(label) for label in relation_labels if str(row[label] or "").strip()}

            # Judged as soon as the row has anything: an x, a value or a reviewer name.
            if not human and not any(str(row[name] or "").strip() for name in (SUPPORTS_NOTHING, "human_value", "reviewer")):
                paper["not_judged"] += 1
                continue
            llm = {name.strip() for name in (compared[row["item_id"]]["llm_relations"] or "").split(";") if name.strip()}
            agreed = human & llm
            value_ok = (compared[row["item_id"]]["value_ok"] or "").strip()
            llm_only = len(llm - human)
            correct = _count(compared[row["item_id"]].get("llm_only_correct"))

            if correct > llm_only:
                warnings.append(f"{path.name} {row['item_id']}: llm_only_correct is {correct}, "
                                f"but there are only {llm_only} LLM-only relations; counted as {llm_only}")
                correct = llm_only

            outcome = Counter(judged=1, exact=int(human == llm), human_none=int(not human),
                              agreed_pairs=len(agreed), llm_pairs=len(llm), human_pairs=len(human),
                              # Corrected: the LLM-only relations the reviewer marked as right count as agreed.
                              llm_only_pairs=llm_only, llm_only_correct=correct,
                              agreed_corrected=len(agreed) + correct, human_corrected=len(human) + correct,
                              exact_corrected=int(human <= llm and llm_only == correct))

            if agreed:
                outcome["value_base"] = 1
                outcome[f"value_{value_ok or 'missing'}"] = 1

            paper += outcome

            for name in llm | human:
                per_relation[name] += Counter(llm=int(name in llm), human=int(name in human), both=int(name in agreed))

        if paper["not_judged"] or paper["value_missing"]:
            warnings.append(f"{path.name}: {paper['not_judged']} quotes not judged, "
                            f"{paper['value_missing']} value_ok cells still empty")

        per_paper[path.name] = _with_rates(paper)
        per_paper[path.name]["quotes_in_paper"] = quotes_per_paper.get(path.name, 0)
        totals += paper

    total = _with_rates(totals)
    total["weighted"] = {
        "exact": _weighted(per_paper, "exact"),
        "precision": _weighted_ratio(per_paper, "agreed_pairs", "llm_pairs"),
        "recall": _weighted_ratio(per_paper, "agreed_pairs", "human_pairs"),
        "values": _weighted_ratio(per_paper, "value_yes", "value_base"),
        "llm_only_right": _weighted_ratio(per_paper, "llm_only_correct", "llm_only_pairs"),
        "exact_corrected": _weighted(per_paper, "exact_corrected"),
        "precision_corrected": _weighted_ratio(per_paper, "agreed_corrected", "llm_pairs"),
        "recall_corrected": _weighted_ratio(per_paper, "agreed_corrected", "human_corrected"),
    }
    per_relation = {
        name: {**counts, "precision": _ratio(counts["both"], counts["llm"]), "recall": _ratio(counts["both"], counts["human"])}
        for name, counts in per_relation.items()
    }

    return {
        "run": str(run_dir),
        "total": total,
        "per_paper": per_paper,
        "per_relation": per_relation,
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
    """Counts plus rates. The 95% Wilson interval is exact for the per-quote measure; for
    precision and recall it is approximate, because pairs from one quote are not independent."""
    result = dict(counts)
    result["exact_rate"] = _wilson(counts["exact"], counts["judged"])
    result["precision"] = _wilson(counts["agreed_pairs"], counts["llm_pairs"])
    result["recall"] = _wilson(counts["agreed_pairs"], counts["human_pairs"])
    result["values"] = _wilson(counts["value_yes"], counts["value_base"])
    result["llm_only_right"] = _wilson(counts["llm_only_correct"], counts["llm_only_pairs"])
    result["exact_corrected_rate"] = _wilson(counts["exact_corrected"], counts["judged"])
    result["precision_corrected"] = _wilson(counts["agreed_corrected"], counts["llm_pairs"])
    result["recall_corrected"] = _wilson(counts["agreed_corrected"], counts["human_corrected"])
    return result


def _count(value) -> int:
    """A whole number typed by a reviewer; empty or unreadable counts as 0."""
    try:
        return max(0, int(float(value)))
    except (TypeError, ValueError):
        return 0


def _ratio(numerator: int, denominator: int) -> float | None:
    return round(numerator / denominator, 3) if denominator else None


def _weighted_ratio(per_paper: dict, numerator: str, denominator: str) -> dict | None:
    """A pooled ratio (such as precision) where each paper's sample counts for all its quotes:
    every judged quote stands for quotes_in_paper / judged quotes of its paper. Point estimate only."""
    top, bottom = 0.0, 0.0

    for paper in per_paper.values():
        if paper.get("judged") and paper.get("quotes_in_paper"):
            scale = paper["quotes_in_paper"] / paper["judged"]
            top += scale * paper.get(numerator, 0)
            bottom += scale * paper.get(denominator, 0)

    return {"rate": round(top / bottom, 3)} if bottom else None


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
    if rate is None:
        return "-"

    if "low" not in rate:
        return f"{rate['rate']:.0%}"

    return f"{rate['rate']:.0%} ({rate['low']:.0%}-{rate['high']:.0%})"


def print_scores(scores: list[dict]) -> None:
    width, column = 46, 24
    print(f"{'':<{width}}" + "".join(f"{Path(s['run']).name:>{column}}" for s in scores))

    for key, label in [("judged", "Quotes judged"), ("not_judged", "Quotes not judged yet"),
                       ("human_none", "Reviewer: quote supports nothing")]:
        print(f"{label:<{width}}" + "".join(f"{s['total'].get(key, 0):>{column}}" for s in scores))

    measures = [("exact", "Exactly the same relations"), ("precision", "Relation precision"),
                ("recall", "Relation recall"), ("values", "Values right (agreed relations)")]
    print("\nEvery paper counts equally (95% CI):")

    for key, label in measures:
        field = "exact_rate" if key == "exact" else key
        print(f"  {label:<{width - 2}}" + "".join(f"{_rate(s['total'].get(field)):>{column}}" for s in scores))

    print("Weighted by each paper's number of quotes (accuracy per quote in the knowledge base):")

    for key, label in measures:
        print(f"  {label:<{width - 2}}" + "".join(f"{_rate(s['total']['weighted'].get(key)):>{column}}" for s in scores))

    # Corrected with llm_only_correct: relations the LLM found and the reviewer missed count as right.
    corrected = [("llm_only_right", "LLM-only relations right after all"),
                 ("exact_corrected", "Exactly the same relations"),
                 ("precision_corrected", "Relation precision"), ("recall_corrected", "Relation recall")]
    print("\nCorrected for relations the reviewer missed, every paper equally (95% CI):")

    for key, label in corrected:
        field = "exact_corrected_rate" if key == "exact_corrected" else key
        print(f"  {label:<{width - 2}}" + "".join(f"{_rate(s['total'].get(field)):>{column}}" for s in scores))

    print("Corrected, weighted by each paper's number of quotes:")

    for key, label in corrected:
        print(f"  {label:<{width - 2}}" + "".join(f"{_rate(s['total']['weighted'].get(key)):>{column}}" for s in scores))

    print(f"\nPer relation (blind, not corrected): precision / recall (times the LLM chose it)")
    relations = dict.fromkeys(relation for s in scores for relation in s["per_relation"])

    for relation in relations:
        cells = []

        for s in scores:
            counts = s["per_relation"].get(relation)
            # "is not None": a precision of exactly 0.0 must still be shown as 0%.
            cells.append("-" if not counts else
                         f"{_rate(None if counts['precision'] is None else {'rate': counts['precision']})} / "
                         f"{_rate(None if counts['recall'] is None else {'rate': counts['recall']})} ({counts['llm']})")

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
