"""Add up the reviewed workbooks of one or more runs into an accuracy estimate.

    python evaluate.py runs/2026-10-17_101500                         # one run
    python evaluate.py runs/2026-10-17_101500 runs/2026-10-20_143000  # compare two runs (round 1, round 2)
    python evaluate.py ... --weighted  # also estimate for all extracted quotes (weights, bootstrap)

Reads every review_*.xlsx (and screening.xlsx, if present) in each run folder,
compares the reviewers' relations per quote (Judge sheet) with the LLM's (Compare
sheet), and prints exact agreement, relation precision and recall, and how often the
LLM's values are right, overall and per relation, plus the include/exclude agreement
of the screening. Saves the result as
scores.json in the run folder. Everything is recomputed from what the reviewers
typed, so it does not depend on the spreadsheet having calculated its formulas.
"""

import argparse
import json
import math
import random
import sys
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import load_workbook

from export import (JUDGE_COLUMNS, JUDGE_REVIEW_COLUMNS, NO_METHOD_FOUND, SUPPORTS_NOTHING, read_jsonl, relation_from_label,
                    review_items, sample_items)
from overview import write_overview

# Each measure is a ratio of two per-quote counts, with the unit it is expressed in.
# Precision and recall count (quote, relation) pairs, not claims: one pair can hold several
# claims (e.g. three activities named in one quote), and the review judges pairs.
MEASURES = {
    "exact": ("exact", "judged", "quotes"),
    "supports_nothing": ("human_none", "judged", "quotes"),
    "precision": ("agreed_pairs", "llm_pairs", "(quote, relation) pairs given by the LLM"),
    "recall": ("agreed_pairs", "human_pairs", "(quote, relation) pairs given by the reviewer, within extracted quotes"),
    "values": ("value_yes", "value_base", "quotes with at least one agreed relation"),
    # Only papers whose Compare step was done; blind and corrected on the same quotes.
    "exact_compared": ("c_exact", "compared", "quotes"),
    "precision_compared": ("c_agreed", "c_llm", "(quote, relation) pairs given by the LLM"),
    "recall_compared": ("c_agreed", "c_human", "(quote, relation) pairs given by the reviewer, within extracted quotes"),
    "llm_only_right": ("llm_only_correct", "llm_only_pairs", "(quote, relation) pairs given only by the LLM"),
    "exact_corrected": ("exact_corrected", "compared", "quotes"),
    "precision_corrected": ("agreed_corrected", "c_llm", "(quote, relation) pairs given by the LLM"),
    "recall_corrected": ("agreed_corrected", "human_corrected",
                         "(quote, relation) pairs given by the reviewer or accepted after comparison"),
}
BOOTSTRAP_SAMPLES = 2000
BOOTSTRAP_SEED = 0


def score_run(run_dir: Path, complete_only: bool = False, weighted: bool = False) -> dict:
    """Agreement between reviewers and LLM for all review workbooks in a run folder.
    With complete_only, papers with quotes not judged yet are left out entirely.
    The results describe the reviewed sample; with weighted, estimates for all extracted
    quotes (sampling-design weights, bootstrap intervals) are added as well.

    Per quote, the reviewer's set of relations is compared with the LLM's:
        exact       both sets are the same
        precision   share of the LLM's (quote, relation) pairs the reviewer also gave
        recall      share of the reviewer's pairs the LLM also gave
        values      share of quotes with an agreed relation whose LLM value is right
    """
    per_paper, relations_by_paper, records, warnings = {}, {}, {}, []
    rows = read_jsonl(run_dir / "claims.jsonl")
    meta_path = run_dir / "meta.json"
    review = json.loads(meta_path.read_text(encoding="utf-8")).get("review", {}) if meta_path.exists() else {}
    publications = {f"review_{Path(publication).stem}.xlsx": publication
                    for publication in dict.fromkeys(row["publication"] for row in rows)}

    for path in sorted(run_dir.glob("review_*.xlsx")):
        workbook = load_workbook(path)
        headers = [cell.value for cell in workbook["Judge"][1]]
        # The relation columns are labelled "relation (concept)".
        relation_labels = headers[len(JUDGE_COLUMNS):len(headers) - len(JUDGE_REVIEW_COLUMNS)]
        compared = {row["item_id"]: row for row in _rows(workbook["Compare"])}
        # The Compare step counts as done once any value_ok or llm_only_correct is filled in.
        compare_done = any(str(row.get(name) or "").strip()
                           for row in compared.values() for name in ("value_ok", "llm_only_correct"))
        paper, paper_relations, paper_records = Counter(), defaultdict(Counter), []
        judge_rows = list(_rows(workbook["Judge"]))
        weights, design = sampling_weights([row for row in rows if row["publication"] == publications.get(path.name)],
                                           [(row["item_id"], row["quote"]) for row in judge_rows], review)

        if not weighted:
            pass  # the weights are only used for the --weighted estimates
        elif not design["matched"]:
            warnings.append(f"{path.name}: the sample could not be reconstructed from claims.jsonl; "
                            "equal weights used within this paper")
        elif design["coverage"] < 1:
            warnings.append(f"{path.name}: {design['quotes_in_paper'] - design['quotes_in_sampled_groups']} of "
                            f"{design['quotes_in_paper']} extracted quotes are in groups that were not sampled; "
                            "their weight is spread over the paper's sampled quotes")

        for row in judge_rows:
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

            # Only filled-in value_ok cells count; an empty one is "not compared yet", not "wrong".
            if agreed:
                outcome[f"value_{value_ok or 'missing'}"] = 1
                outcome["value_base"] = int(bool(value_ok))

            # The corrected measures only use papers whose Compare step was done; otherwise
            # an empty llm_only_correct would be read as "0 overlooked relations". For those
            # papers, c_* keep the blind counts too, so blind and corrected can be compared
            # on exactly the same quotes.
            if not compare_done:
                for key in ("llm_only_pairs", "llm_only_correct", "agreed_corrected", "human_corrected", "exact_corrected"):
                    del outcome[key]
            else:
                outcome.update(compared=1, c_exact=int(human == llm), c_agreed=len(agreed),
                               c_llm=len(llm), c_human=len(human))

            paper += outcome
            paper_records.append((weights.get(row["item_id"], 1.0), outcome))

            for name in llm | human:
                paper_relations[name] += Counter(llm=int(name in llm), human=int(name in human), both=int(name in agreed))

        if not compare_done and paper["judged"]:
            warnings.append(f"{path.name}: Compare step not done yet; left out of values and corrected measures")
        elif paper["not_judged"] or paper["value_missing"]:
            warnings.append(f"{path.name}: {paper['not_judged']} quotes not judged, "
                            f"{paper['value_missing']} value_ok cells still empty")

        if complete_only and (paper["not_judged"] or not paper["judged"]):
            warnings.append(f"{path.name}: left out (--complete-only): {paper['not_judged']} quotes not judged")
            continue

        per_paper[path.name] = {**_with_rates(paper), **design}
        per_paper[path.name]["counts"] = paper
        relations_by_paper[path.name] = paper_relations
        records[path.name] = paper_records

    totals = sum((paper.pop("counts") for paper in per_paper.values()), Counter())
    per_relation = defaultdict(Counter)

    for paper_relations in relations_by_paper.values():
        for name, counts in paper_relations.items():
            per_relation[name] += counts

    total = _with_rates(totals)
    # The sample itself: every measure as numerator / denominator over the reviewed quotes.
    total["sample"] = {name: {"rate": _ratio(totals[num], totals[den]), "num": totals[num], "den": totals[den],
                              "unit": unit}
                       for name, (num, den, unit) in MEASURES.items()}
    # Estimates for all extracted quotes of the reviewed papers, from the sample (only with --weighted).
    total["estimate"] = design_estimates({name: records[name] for name in per_paper}) if weighted else None
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


def reference_decision(c1, c2, c3) -> str | None:
    """The reference decision for a method, derived from the reviewer's judgement of C1-C3
    (met / not met / not in this paper); None when not all three are judged.

        any criterion not met                          exclude
        C1 or C2 not in this paper                     exclude: the paper does not describe the
                                                       method well enough to extract it
        all met                                        include
        C1 and C2 met, C3 not in this paper            pending: the ICT application has to come
                                                       from another publication
    """
    judgements = [str(value or "").strip() for value in (c1, c2, c3)]

    if not all(judgements):
        return None
    if "not met" in judgements or "not in this paper" in judgements[:2]:
        return "exclude"
    return "include" if judgements[2] == "met" else "pending"


def score_screening(path: Path, warnings: list[str]) -> dict | None:
    """Include/exclude agreement from screening.xlsx (if the run has one), per method and per paper.

    Rows the reviewer added on the Judge sheet are methods the LLM did not find:
    marked "include", they count as missed by the LLM.
    """
    if not path.exists():
        return None

    workbook = load_workbook(path)
    llm = {(row["paper"], row["method"]): row for row in _rows(workbook["Compare"])}
    methods = Counter()
    criteria = defaultdict(Counter)
    papers = defaultdict(lambda: {"human": False, "llm": False, "judged": False})

    for (paper, _), row in llm.items():
        papers[paper]["llm"] |= row["llm_decision"] == "include"

    for row in _rows(workbook["Judge"]):
        # The reference decision follows from the reviewer's criterion judgements (reference_decision);
        # only rows without criteria, such as "(no method found)", keep the reviewer's typed decision.
        human = reference_decision(*(row.get(criterion) for criterion in ("C1", "C2", "C3"))) \
            or str(row.get("human_decision") or "").strip()
        key = (row["paper"], row["method"])
        typed = str(row.get("human_decision") or "").strip()

        if typed and human != typed:
            methods["decision_derived_differs"] += 1

        if not human:
            methods["not_judged"] += 1
            continue

        papers[row["paper"]]["judged"] = True
        papers[row["paper"]]["human"] |= human == "include"

        if key not in llm:  # added by the reviewer
            methods["missed_by_llm"] += human == "include"
            continue

        # A paper without candidate methods is a paper-level result, not a method.
        if row["method"] == NO_METHOD_FOUND:
            methods["papers_without_candidates"] += 1
            methods["papers_without_candidates_human_include"] += human == "include"
            continue

        llm_decision = llm[key]["llm_decision"]
        methods["judged"] += 1
        methods[f"llm_{llm_decision}_human_{human}"] += 1

        # "pending" (evidence must come from another paper) is a third outcome, not a disagreement:
        # the decision agreement is computed over methods the reviewer included or excluded.
        if human == "pending":
            methods["pending"] += 1
        else:
            methods["decided"] += 1
            methods["same_decision"] += human == llm_decision

        # Per criterion: agreement where both say met / not met; "not in this paper" counted apart.
        for criterion in ("C1", "C2", "C3"):
            human_value = str(row.get(criterion) or "").strip()
            llm_value = str(llm[key].get(f"llm_{criterion}") or "").strip()

            if human_value == "not in this paper":
                criteria[criterion]["not_in_paper"] += 1
            elif human_value and llm_value:
                criteria[criterion]["judged"] += 1
                criteria[criterion]["same"] += human_value == llm_value

    if methods["not_judged"]:
        warnings.append(f"screening.xlsx: {methods['not_judged']} methods not judged")

    paper_counts = Counter()

    for paper in papers.values():
        if paper["judged"]:
            paper_counts["judged"] += 1
            paper_counts["same_decision"] += paper["human"] == paper["llm"]
            paper_counts[f"llm_{'include' if paper['llm'] else 'exclude'}_human_{'include' if paper['human'] else 'exclude'}"] += 1

    methods = dict(methods)
    methods["same_decision_rate"] = _wilson(methods.get("same_decision", 0), methods.get("decided", 0))
    paper_counts = dict(paper_counts)
    paper_counts["same_decision_rate"] = _wilson(paper_counts.get("same_decision", 0), paper_counts.get("judged", 0))
    per_criterion = {criterion: {**counts, "same_rate": _wilson(counts["same"], counts["judged"])}
                     for criterion, counts in criteria.items()}
    return {"per_method": methods, "per_paper": paper_counts, "per_criterion": per_criterion}


def _with_rates(counts: Counter) -> dict:
    """Counts plus rates. The 95% Wilson interval is exact for the per-quote measure; for
    precision and recall it is approximate, because pairs from one quote are not independent."""
    result = dict(counts)
    result["exact_rate"] = _wilson(counts["exact"], counts["judged"])
    result["precision"] = _wilson(counts["agreed_pairs"], counts["llm_pairs"])
    result["recall"] = _wilson(counts["agreed_pairs"], counts["human_pairs"])
    result["values"] = _wilson(counts["value_yes"], counts["value_base"])
    result["llm_only_right"] = _wilson(counts["llm_only_correct"], counts["llm_only_pairs"])
    result["exact_corrected_rate"] = _wilson(counts["exact_corrected"], counts["compared"])
    result["precision_corrected"] = _wilson(counts["agreed_corrected"], counts["c_llm"])
    result["recall_corrected"] = _wilson(counts["agreed_corrected"], counts["human_corrected"])
    # Blind, on the same papers as the corrected measures.
    result["exact_compared"] = _wilson(counts["c_exact"], counts["compared"])
    result["precision_compared"] = _wilson(counts["c_agreed"], counts["c_llm"])
    result["recall_compared"] = _wilson(counts["c_agreed"], counts["c_human"])
    return result


def _count(value) -> int:
    """A whole number typed by a reviewer; empty or unreadable counts as 0."""
    try:
        return max(0, int(float(value)))
    except (TypeError, ValueError):
        return 0


def _ratio(numerator: int, denominator: int) -> float | None:
    return round(numerator / denominator, 3) if denominator else None


def sampling_weights(paper_rows: list[dict], sampled: list[tuple[str, str]], review: dict) -> tuple[dict, dict]:
    """Weight of each sampled quote: the number of extracted quotes it stands for.

    The sample (export.sample_items) is stratified: quotes are grouped by (method, first
    relation), and each group gets about the same number of places, so a quote from a small
    group is more likely to be sampled than one from a large group. A sampled quote from a
    group of N_g quotes of which n_g were sampled stands for N_g / n_g quotes (inverse
    inclusion probability). Groups with no sampled quote at all (more groups than places)
    are covered by scaling the paper's weights up to its total number of quotes, i.e. by
    assuming they resemble the rest of the paper.

    The sample is reconstructed from claims.jsonl with the run's sample size and seed, and
    checked against the quotes in the workbook.
    """
    items = review_items(paper_rows)
    groups = Counter((item["method"], item["claims"][0]["relation"]) for item in items)
    sample = sample_items(review_items(paper_rows), review.get("sample_per_paper", 20), review.get("seed", 1))
    design = {"quotes_in_paper": len(items), "groups": len(groups)}
    matched = len(sample) == len(sampled) and all(item["quote"] == quote for item, (_, quote) in zip(sample, sampled))

    if not matched or not sample:
        size = len(items) or len(sampled)
        return ({item_id: size / len(sampled) for item_id, _ in sampled} if sampled else {},
                {**design, "matched": False, "coverage": None, "quotes_in_sampled_groups": None})

    sampled_per_group = Counter((item["method"], item["claims"][0]["relation"]) for item in sample)
    raw = [groups[key] / sampled_per_group[key]
           for key in ((item["method"], item["claims"][0]["relation"]) for item in sample)]
    in_sampled_groups = sum(groups[key] for key in sampled_per_group)
    scale = len(items) / sum(raw)
    weights = {item_id: weight * scale for (item_id, _), weight in zip(sampled, raw)}
    return weights, {**design, "matched": True, "groups_sampled": len(sampled_per_group),
                     "quotes_in_sampled_groups": in_sampled_groups, "coverage": round(in_sampled_groups / len(items), 3)}


def design_estimates(records: dict[str, list[tuple[float, Counter]]]) -> dict:
    """Every measure as a weighted ratio over all papers (sum of w * numerator over sum of
    w * denominator), with a 95% bootstrap interval: quotes are resampled with replacement
    within each paper (keeping their weights), and the 2.5th and 97.5th percentiles are taken.
    Resampling whole quotes keeps the relations of one quote together, which matters for
    precision and recall."""
    papers = [paper for paper in records.values() if paper]
    rng = random.Random(BOOTSTRAP_SEED)
    # Each count once, even when several measures use it (agreed_pairs is in precision and recall).
    keys = {key for numerator, denominator, _ in MEASURES.values() for key in (numerator, denominator)}

    def estimate(sample) -> dict:
        sums = defaultdict(float)

        for weight, outcome in sample:
            for key in keys:
                sums[key] += weight * outcome.get(key, 0)

        return {name: (sums[num] / sums[den] if sums[den] else None) for name, (num, den, _) in MEASURES.items()}

    point = estimate([record for paper in papers for record in paper])
    replicates = defaultdict(list)

    for _ in range(BOOTSTRAP_SAMPLES):
        resample = [record for paper in papers for record in rng.choices(paper, k=len(paper))]

        for name, value in estimate(resample).items():
            if value is not None:
                replicates[name].append(value)

    result = {}

    for name, (numerator, denominator, unit) in MEASURES.items():
        if point[name] is None:
            result[name] = None
            continue

        values = sorted(replicates[name])
        result[name] = {"rate": round(point[name], 3),
                        "low": round(values[int(0.025 * (len(values) - 1))], 3),
                        "high": round(values[int(0.975 * (len(values) - 1))], 3),
                        "unit": unit,
                        "n": sum(outcome.get(denominator, 0) for paper in papers for _, outcome in paper)}

    return result


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

    sections = [
        ("Blind", [("exact", "Exactly the same relations"), ("precision", "Relation precision"),
                   ("recall", "Relation recall (within extracted quotes)"),
                   ("values", "Values right"), ("supports_nothing", "Quote supports nothing")]),
        ("Blind, papers with Compare done (for comparison with corrected)",
         [("exact_compared", "Exactly the same relations"), ("precision_compared", "Relation precision"),
          ("recall_compared", "Relation recall (within extracted quotes)")]),
        ("Corrected for relations the reviewer missed",
         [("llm_only_right", "LLM-only relations right after all"), ("exact_corrected", "Exactly the same relations"),
          ("precision_corrected", "Relation precision"), ("recall_corrected", "Relation recall (within extracted quotes)")]),
    ]

    def sample_cell(s: dict, key: str) -> str:
        m = s["total"]["sample"][key]
        return "-" if not m["den"] else f"{_whole_percent(m['num'], m['den'])}% ({m['num']}/{m['den']})"

    print("\nIn the reviewed sample (descriptive, unweighted):")

    for title, measures in sections:
        print(f"  {title}:")

        for key, label in measures:
            print(f"    {label:<{width - 4}}" + "".join(f"{sample_cell(s, key):>{column}}" for s in scores))

    print("\n  Units: " + "; ".join(f"{name} per {MEASURES[name][2]}" for name in ("exact", "precision", "recall", "values")))
    print("  Values: " + "".join(f"{s['total'].get('value_yes', 0)} yes / {s['total'].get('value_partly', 0)} partly / "
                                 f"{s['total'].get('value_no', 0)} no   " for s in scores))

    if any(s["total"]["estimate"] for s in scores):
        print("\nEstimates for all extracted quotes of the reviewed papers "
              "(--weighted: sampling-design weights, 95% bootstrap interval):")

        for title, measures in sections:
            print(f"  {title}:")

            for key, label in measures:
                print(f"    {label:<{width - 4}}" + "".join(
                    f"{_rate((s['total']['estimate'] or {}).get(key)):>{column}}" for s in scores))

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
        method_extra = [("pending", "Reviewer: pending (evidence needed from another paper)"),
                        ("llm_exclude_human_pending", "  of which the LLM excluded"),
                        ("missed_by_llm", "Missed by the LLM (added by reviewer)")]

        for level, title, extra in [("per_method", "Screening per method (same decision over included/excluded)", method_extra),
                                    ("per_paper", "Screening per paper (included if at least one method is)", [])]:
            print(f"\n{title}")

            for key, label in [("judged", "Judged"),
                               ("llm_include_human_exclude", "LLM included, reviewer excluded"),
                               ("llm_exclude_human_include", "LLM excluded, reviewer included")] + extra:
                print(f"  {label:<{width - 2}}" + "".join(
                    f"{s['screening'][level].get(key, 0) if s['screening'] else '-':>{column}}" for s in scores))

            def same(s: dict) -> str:
                if not s["screening"]:
                    return "-"
                counts = s["screening"][level]
                base = counts.get("decided", counts.get("judged", 0)) if level == "per_method" else counts.get("judged", 0)
                return f"{counts.get('same_decision', 0)}/{base}"

            print(f"  {'Same decision':<{width - 2}}" + "".join(f"{same(s):>{column}}" for s in scores))

        print("\nScreening per criterion (same judgement where both say met / not met)")

        for criterion in ("C1", "C2", "C3"):
            cells = []

            for s in scores:
                counts = (s["screening"] or {}).get("per_criterion", {}).get(criterion)
                cells.append("-" if not counts else
                             f"same {counts.get('same', 0)}/{counts.get('judged', 0)}, "
                             f"not in paper {counts.get('not_in_paper', 0)}")

            print(f"  {criterion:<{width - 2}}" + "".join(f"{cell:>{column + 14}}" for cell in cells))

    for s in scores:
        for warning in s["warnings"]:
            print(f"\nWarning ({Path(s['run']).name}): {warning}")


# The measures written to LaTeX, with their macro names and labels for the comparison table.
LATEX_MEASURES = [
    ("exact", "Exact", "Exact agreement"),
    ("supports_nothing", "SupportsNothing", "Quote supports nothing"),
    ("precision", "Precision", "Relation precision"),
    ("recall", "Recall", "Relation recall (within extracted quotes)"),
    ("values", "Values", "Value correctness"),
    ("llm_only_right", "LlmOnlyRight", "Pipeline-only relations accepted"),
    ("exact_corrected", "ExactCorrected", "Exact agreement, corrected"),
    ("precision_corrected", "PrecisionCorrected", "Relation precision, corrected"),
    ("recall_corrected", "RecallCorrected", "Relation recall, corrected"),
]
# Macro prefixes when several runs are evaluated together (LaTeX names cannot contain digits).
ROUND_PREFIXES = ["ResOne", "ResTwo", "ResThree", "ResFour"]


def _whole_percent(numerator: int, denominator: int) -> int:
    """Percentage rounded half up, computed exactly (39/120 = 32.5% gives 33, not 32)."""
    return (200 * numerator + denominator) // (2 * denominator)


def _percent_count(measure: dict) -> tuple[str, str]:
    """A sample measure as LaTeX: ("54\\%", "89/166")."""
    if not measure["den"]:
        return "--", "--"
    return f"{_whole_percent(measure['num'], measure['den'])}\\%", f"{measure['num']}/{measure['den']}"


def write_latex(score: dict, path: Path, prefix: str = "Res") -> None:
    """The headline numbers as LaTeX commands, plus a table per paper, for the paper.

    \\input{results.tex} in the paper and write e.g. "precision was \\ResPrecision{}
    (\\ResPrecisionN{})", which gives "54% (89/166)"; re-running evaluate.py then updates
    every number in the paper. The numbers describe the reviewed sample; the weighted
    estimates are only added (as \\ResWeighted...) when evaluate.py ran with --weighted.
    """
    total, estimate = score["total"], score["total"]["estimate"]

    def percent(rate: dict | None) -> str:
        return "--" if rate is None else f"{rate['rate'] * 100:.0f}\\%"

    def interval(rate: dict | None) -> str:
        return "--" if rate is None or "low" not in rate else f"{rate['low'] * 100:.0f}--{rate['high'] * 100:.0f}\\%"

    def ratio(numerator: int, denominator: int) -> str:
        return f"{numerator / denominator * 100:.0f}\\%" if denominator else "--"

    commands = {
        # Counts
        "Papers": str(len(score["per_paper"])),
        "PapersCompared": str(sum(1 for paper in score["per_paper"].values() if paper.get("compared"))),
        "Quotes": str(total.get("judged", 0)),
        "QuotesCompared": str(total.get("compared", 0)),
        "PairsLlm": str(total.get("llm_pairs", 0)),
        "PairsHuman": str(total.get("human_pairs", 0)),
        "PairsAgreed": str(total.get("agreed_pairs", 0)),
        "QuotesInPapers": str(sum(paper.get("quotes_in_paper", 0) for paper in score["per_paper"].values())),
        "ValuesYes": str(total.get("value_yes", 0)),
        "ValuesPartly": str(total.get("value_partly", 0)),
        "ValuesWrong": str(total.get("value_no", 0)),
    }

    # The sample (descriptive): \ResPrecision{} = "54\%" and \ResPrecisionN{} = "89/166", ...
    for key, name, _ in LATEX_MEASURES:
        commands[name], commands[f"{name}N"] = _percent_count(total["sample"][key])

    # Only with --weighted: estimates for all extracted quotes with their 95% bootstrap interval.
    if estimate:
        for key, name, _ in LATEX_MEASURES:
            commands[f"Weighted{name}"] = percent(estimate.get(key))
            commands[f"Weighted{name}CI"] = interval(estimate.get(key))

    # Screening per method and per criterion (counts), if the run has a screening sheet.
    screening = score.get("screening")
    if screening:
        methods = screening["per_method"]
        commands.update({
            "ScreenMethodsJudged": str(methods.get("judged", 0)),
            "ScreenMethodsDecided": str(methods.get("decided", 0)),
            "ScreenMethodsSame": str(methods.get("same_decision", 0)),
            "ScreenLlmInclHumanExcl": str(methods.get("llm_include_human_exclude", 0)),
            "ScreenLlmExclHumanIncl": str(methods.get("llm_exclude_human_include", 0)),
            "ScreenPending": str(methods.get("pending", 0)),
            "ScreenPendingLlmExcl": str(methods.get("llm_exclude_human_pending", 0)),
            "ScreenPapersWithoutCandidates": str(methods.get("papers_without_candidates", 0)),
        })
        for criterion, word in (("C1", "One"), ("C2", "Two"), ("C3", "Three")):
            counts = screening["per_criterion"].get(criterion, {})
            commands[f"ScreenC{word}Same"] = str(counts.get("same", 0))
            commands[f"ScreenC{word}Judged"] = str(counts.get("judged", 0))
            commands[f"ScreenC{word}NotInPaper"] = str(counts.get("not_in_paper", 0))

    lines = ["% Generated by evaluate.py from " + score["run"] + " -- do not edit by hand."]
    lines += [f"\\newcommand{{\\{prefix}{name}}}{{{value}}}" for name, value in commands.items()]
    # Per paper: the sample itself, unweighted (rates over the sampled quotes and their pairs).
    lines += [
        "",
        f"\\newcommand{{\\{prefix}PaperTable}}{{%",
        "\\begin{tabular}{lrrrrrrr}",
        "\\hline",
        "Publication & Extracted & Sampled & Exact & Precision & Recall & Values & Corr.\\ precision \\\\",
        "\\hline",
    ]

    for name, paper in score["per_paper"].items():
        title = name.removeprefix("review_").removesuffix(".xlsx").replace("_", " ")
        corrected = ratio(paper.get("agreed_corrected", 0), paper.get("c_llm", 0)) if paper.get("compared") else "--"
        lines.append(f"{title} & {paper.get('quotes_in_paper', 0)} & {paper.get('judged', 0)} & "
                     f"{ratio(paper.get('exact', 0), paper.get('judged', 0))} & "
                     f"{ratio(paper.get('agreed_pairs', 0), paper.get('llm_pairs', 0))} & "
                     f"{ratio(paper.get('agreed_pairs', 0), paper.get('human_pairs', 0))} & "
                     f"{ratio(paper.get('value_yes', 0), paper.get('value_base', 0))} & {corrected} \\\\")

    lines += ["\\hline", "\\end{tabular}}", ""]
    path.write_text("\n".join(lines), encoding="utf-8")


def comparison_latex(scores: list[dict]) -> str:
    """\\ResComparisonTable: every measure per round, as "54\\% (89/166)"."""
    lines = ["% Comparison of the evaluated runs, in the order given to evaluate.py.",
             "\\newcommand{\\ResComparisonTable}{%",
             "\\begin{tabular}{l" + "r" * len(scores) + "}",
             "\\hline",
             "Measure & " + " & ".join(f"Round {i + 1}" for i in range(len(scores))) + " \\\\",
             "\\hline"]

    for key, _, label in LATEX_MEASURES:
        cells = []
        for score in scores:
            rate, count = _percent_count(score["total"]["sample"][key])
            cells.append("--" if rate == "--" else f"{rate} ({count})")
        lines.append(f"{label} & " + " & ".join(cells) + " \\\\")

    lines += ["\\hline", "\\end{tabular}}", ""]
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("run_dirs", nargs="+", type=Path, help="run folders with reviewed workbooks")
    parser.add_argument("--complete-only", action="store_true",
                        help="leave out papers whose review is not finished")
    parser.add_argument("--weighted", action="store_true",
                        help="also estimate the results for all extracted quotes (sampling-design weights, "
                             "bootstrap intervals); by default the reviewed sample is reported descriptively")
    args = parser.parse_args()
    scores = []

    if len(args.run_dirs) > len(ROUND_PREFIXES):
        sys.exit(f"At most {len(ROUND_PREFIXES)} runs at once")

    # One run: macros \Res...; several: \ResOne..., \ResTwo..., so the paper can \input all of them.
    prefixes = ["Res"] if len(args.run_dirs) == 1 else ROUND_PREFIXES[:len(args.run_dirs)]

    for run_dir in args.run_dirs:
        if not list(run_dir.glob("review_*.xlsx")) and not (run_dir / "screening.xlsx").exists():
            sys.exit(f"No review_*.xlsx or screening.xlsx in {run_dir}")

        scores.append(score_run(run_dir, args.complete_only, args.weighted))

    for run_dir, result, prefix in zip(args.run_dirs, scores, prefixes):
        (run_dir / "scores.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
        write_latex(result, run_dir / "results.tex", prefix)
        write_overview(run_dir, run_dir / "review_overview.md")

    # The comparison table goes into the last run's results.tex.
    if len(scores) > 1:
        with open(args.run_dirs[-1] / "results.tex", "a", encoding="utf-8") as file:
            file.write("\n" + comparison_latex(scores))

    print_scores(scores)
    for run_dir, prefix in zip(args.run_dirs, prefixes):
        print(f"\nWritten to {run_dir}: results.tex (LaTeX macros \\{prefix}...), "
              "review_overview.md (raw counts and notes), scores.json")
    if len(scores) > 1:
        print(f"Comparison table \\ResComparisonTable: in {args.run_dirs[-1] / 'results.tex'}")


if __name__ == "__main__":
    main()
