"""Run the whole pipeline on one or more PDFs, and save everything needed to trace the result.

    python run.py                                   # every PDF in papers/dev/
    python run.py --final                           # every PDF in papers/heldout/ (final evaluation)
    python run.py papers/dev/paper.pdf --method "Value Sensitive Design"
    python run.py --rebuild runs/2026-10-11_153000  # redo verify + export, no API calls
    python run.py --retry-failed runs/2026-10-11_153000  # redo only the papers that failed

Each run gets its own folder under runs/ containing:
    meta.json        what was run: inputs, ontology, prompts, models, settings, timings
    extracted.jsonl  the model's raw claims, saved as soon as each method is done
    claims.jsonl     the claims after verification (input for the knowledge graph)
    review_<paper>.xlsx  one workbook per paper for human review (see export.py)
"""

import argparse
import hashlib
import json
import os
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from importlib.metadata import version
from pathlib import Path
from threading import Lock

from dotenv import load_dotenv
from mistralai.client import Mistral
from mistralai.client.utils import BackoffStrategy, RetryConfig

import extract
from export import read_jsonl, write_jsonl, write_review_sheet, write_screening_sheet
from ocr import CACHE_DIR, OCR_MODEL, ocr_pdf
from ontology import load_ontology
from verify import verify_claims

ONTOLOGY_PATH = Path("config/ontology.yaml")
RUNS_DIR = Path("runs")
HELDOUT_FOLDER = "heldout"  # any PDF inside a folder with this name counts as held out
DEV_PAPERS = Path("papers/dev")  # processed when no PDFs are given
HELDOUT_PAPERS = Path("papers/heldout")  # processed when no PDFs are given, with --final
REVIEW_SAMPLE = 20  # quotes per paper that reviewers judge
REVIEW_SEED = 1  # fixed, so the same run always gives the same review sample
# The SDK does not retry by default, so one rate limit (429), server error (5xx), dropped
# connection or timeout would fail a whole paper. Retry with growing waits (2 s up to
# 1 min) for at most 45 minutes per call, which leaves room for one retry after a timeout.
RETRIES = RetryConfig("backoff", BackoffStrategy(2_000, 60_000, 2.0, 45 * 60_000), retry_connection_errors=True)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("pdfs", nargs="*", type=Path,
                        help=f"PDF files to process; default: all PDFs in {DEV_PAPERS}/ (or {HELDOUT_PAPERS}/ with --final)")
    parser.add_argument("--method", action="append", default=[],
                        help="method to extract, for every PDF (repeat for several); default: the model identifies them")
    parser.add_argument("--workers", type=int, default=3, help="number of papers processed at the same time")
    parser.add_argument("--final", action="store_true", help="allow held-out papers (final evaluation only)")
    parser.add_argument("--rebuild", type=Path, metavar="RUN_DIR",
                        help="verify and export an earlier run again from its extracted.jsonl, without API calls")
    parser.add_argument("--retry-failed", type=Path, metavar="RUN_DIR",
                        help="process the papers that failed in an earlier run again, into that same run")
    parser.add_argument("--sample", type=int, default=None,
                        help=f"quotes per paper in the review sheet (default {REVIEW_SAMPLE}, 0 = all)")
    parser.add_argument("--seed", type=int, default=None,
                        help=f"random seed for picking the review sample (default {REVIEW_SEED})")
    args = parser.parse_args()

    if args.rebuild:
        rebuild(args.rebuild, args.sample, args.seed)
        return

    if args.retry_failed:
        retry_failed(args.retry_failed, args.workers)
        return

    if not args.pdfs:
        folder = HELDOUT_PAPERS if args.final else DEV_PAPERS
        args.pdfs = sorted(folder.glob("*.pdf"))

        if not args.pdfs:
            sys.exit(f"No PDFs in {folder}/")

        print(f"Processing all {len(args.pdfs)} PDFs in {folder}/")

    missing = [pdf for pdf in args.pdfs if not pdf.is_file()]

    if missing:
        sys.exit("Not found: " + ", ".join(map(str, missing)))

    check_heldout(args.pdfs, args.final)
    review = {"sample_per_paper": REVIEW_SAMPLE if args.sample is None else args.sample,
              "seed": REVIEW_SEED if args.seed is None else args.seed}
    run(args.pdfs, args.method, args.workers, args.final, review)


def check_heldout(pdfs: list[Path], final: bool) -> None:
    """Held-out papers are only for the final evaluation, never for prompt development."""
    heldout = [pdf for pdf in pdfs if HELDOUT_FOLDER in pdf.resolve().parts]
    
    if heldout and not final:
        sys.exit("Refusing to run on held-out papers without --final:\n" + "\n".join(f"  {pdf}" for pdf in heldout))


def run(pdfs: list[Path], methods: list[str], workers: int, final: bool, review: dict) -> None:
    load_dotenv()
    client = Mistral(api_key=os.environ["MISTRAL_API_KEY"], retry_config=RETRIES)
    ontology = load_ontology(ONTOLOGY_PATH)

    run_dir = RUNS_DIR / datetime.now().strftime("%Y-%m-%d_%H%M%S")
    run_dir.mkdir(parents=True)
    raw_file = run_dir / "extracted.jsonl"
    meta = {
        "started": datetime.now().isoformat(timespec="seconds"),
        "command": " ".join(sys.argv),
        "final_evaluation": final,
        "git_commit": git_commit(),
        **settings(ontology),
        "mistralai_version": version("mistralai"),
        "methods_given": methods,
        "review": review,
        "papers": [],
    }
    save_meta(run_dir, meta)  # saved now too, so even a crashed run leaves a trace
    print(f"Run folder: {run_dir}\n")

    lock = Lock()  # papers run in parallel threads, but share one extracted.jsonl

    def process(pdf: Path) -> tuple[dict, list[dict]]:
        return process_paper(client, pdf, methods, ontology, raw_file, lock)

    with ThreadPoolExecutor(max_workers=workers) as pool:
        results = list(pool.map(process, pdfs))  # map() returns results in input order

    meta["papers"] = [paper for paper, _ in results]
    rows = [row for _, paper_rows in results for row in paper_rows]
    finish(run_dir, rows, meta, ontology)


def settings(ontology: dict) -> dict:
    """Everything that determines the extraction results. Recorded in meta.json, and it
    must not change within a run, so the papers of one run are comparable."""
    return {
        "ontology": {"path": str(ONTOLOGY_PATH), "version": ontology["version"], "sha256": ontology["hash"]},
        # Hash of the prompt text the run actually uses (loaded once, see extract.load_prompt).
        "prompts": {name: hashlib.sha256(extract.load_prompt(name).encode("utf-8")).hexdigest()[:12]
                    for name in ("identify", "extract")},
        "extraction": {
            "model": extract.EXTRACTION_MODEL,
            "temperature": extract.TEMPERATURE,
            "top_p": extract.TOP_P,
            "random_seed": extract.RANDOM_SEED,
            "reasoning_effort": extract.REASONING_EFFORT,
        },
        "ocr_model": OCR_MODEL,
    }


def retry_failed(run_dir: Path, workers: int) -> None:
    """Process the papers that failed in an earlier run again, into the same run folder."""
    meta = json.loads((run_dir / "meta.json").read_text(encoding="utf-8"))
    failed = [index for index, paper in enumerate(meta["papers"]) if paper["error"]]

    if not failed:
        sys.exit(f"No failed papers in {run_dir}")

    ontology = load_ontology(ONTOLOGY_PATH)
    # Same settings as the original run, otherwise one run would mix two versions.
    changed = [key for key, value in settings(ontology).items() if meta.get(key) != value]

    if changed:
        sys.exit(f"Settings changed since this run ({', '.join(changed)}). "
                 "Start a new run instead, or undo the changes.")

    # Drop the claims these papers saved before they failed; they are extracted again.
    failed_shas = {meta["papers"][index]["sha256"] for index in failed} - {None}
    raw_file = run_dir / "extracted.jsonl"
    write_jsonl([row for row in read_jsonl(raw_file) if row["pdf_sha256"] not in failed_shas], raw_file)

    load_dotenv()
    client = Mistral(api_key=os.environ["MISTRAL_API_KEY"], retry_config=RETRIES)
    lock = Lock()
    print(f"Retrying {len(failed)} failed papers in {run_dir}\n")

    def process(index: int) -> tuple[dict, list[dict]]:
        pdf = Path(meta["papers"][index]["pdf"])
        return process_paper(client, pdf, meta["methods_given"], ontology, raw_file, lock)

    with ThreadPoolExecutor(max_workers=workers) as pool:
        results = list(pool.map(process, failed))

    for index, (paper, _) in zip(failed, results):
        meta["papers"][index] = paper

    meta.setdefault("retried", []).append({
        "at": datetime.now().isoformat(timespec="seconds"),
        "git_commit": git_commit(),
        "papers": [meta["papers"][index]["pdf"] for index in failed],
    })
    save_meta(run_dir, meta)
    rebuild(run_dir, None, None)  # verify and export the whole run again, now with these papers


def process_paper(client: Mistral, pdf: Path, methods: list[str], ontology: dict,
                  raw_file: Path, lock: Lock) -> tuple[dict, list[dict]]:
    """OCR, extraction and verification for one PDF. Errors are recorded, not raised,
    so one bad paper does not stop the others."""
    started = time.time()
    paper = {"pdf": str(pdf), "sha256": None, "pages": None, "ocr_model": None,
             "screening": None, "methods": [], "error": None}
    rows = []

    try:
        ocr = ocr_pdf(client, pdf)
        paper.update(sha256=ocr["sha256"], pages=len(ocr["pages"]), ocr_model=ocr["model"])
        print(f"{pdf.name}: OCR ready ({len(ocr['pages'])} pages)")

        if methods:
            paper["methods"] = methods  # chosen by hand, so the paper is not screened
        else:
            paper["screening"] = extract.screen_methods(client, ocr, ontology)
            paper["methods"] = [candidate["name"] for candidate in paper["screening"] if candidate["included"]]

            for candidate in paper["screening"]:
                failed = [criterion for criterion, check in candidate["criteria"].items() if not check["met"]]
                print(f"{pdf.name}: {candidate['name']}: "
                      + ("included" if candidate["included"] else f"rejected ({', '.join(failed)} not met)"))

            if not paper["methods"]:
                print(f"{pdf.name}: paper rejected, no method meets the criteria")

        for method in paper["methods"]:
            method_rows = extract.extract_claims(client, ocr, method, ontology)
            
            for row in method_rows:
                # Unique within a run, and the same again after --rebuild.
                row["claim_id"] = f"{ocr['sha256'][:6]}-{len(rows) + 1:03d}"
                rows.append(row)
            
            with lock:
                with open(raw_file, "a", encoding="utf-8") as file:
                    for row in method_rows:
                        file.write(json.dumps(row, ensure_ascii=False) + "\n")
            
            print(f"{pdf.name}: {method}: {sum(r['status'] == 'reported' for r in method_rows)} claims")

        verify_claims(rows, ocr)
    
    except Exception as error:
        paper["error"] = f"{type(error).__name__}: {error}"
        print(f"{pdf.name}: FAILED: {paper['error']}  (claims saved so far are in {raw_file.name})")
        rows = []
    
    paper["seconds"] = round(time.time() - started)
    paper["claims"] = sum(row["status"] == "reported" for row in rows)
    
    return paper, rows


def rebuild(run_dir: Path, sample: int | None, seed: int | None) -> None:
    """Verify and export an earlier run again, e.g. after changing verify.py or export.py.
    The review sample is drawn the same way as before unless --sample or --seed is given."""
    meta = json.loads((run_dir / "meta.json").read_text(encoding="utf-8"))
    review = meta.get("review", {"sample_per_paper": REVIEW_SAMPLE, "seed": REVIEW_SEED})

    if sample is not None:
        review["sample_per_paper"] = sample

    if seed is not None:
        review["seed"] = seed

    meta["review"] = review
    rows = read_jsonl(run_dir / "extracted.jsonl")
    # Restore the paper order of the original run (threads may have saved them in another order).
    paper_order = [paper["sha256"] for paper in meta["papers"]]
    rows.sort(key=lambda row: (paper_order.index(row["pdf_sha256"]) if row["pdf_sha256"] in paper_order else 0,
                               row["claim_id"]))
    
    for sha256 in dict.fromkeys(row["pdf_sha256"] for row in rows):
        ocr = json.loads((CACHE_DIR / f"{sha256}.json").read_text(encoding="utf-8"))
        verify_claims([row for row in rows if row["pdf_sha256"] == sha256], ocr)
    
    meta.setdefault("rebuilt", []).append(datetime.now().isoformat(timespec="seconds"))
    finish(run_dir, rows, meta, load_ontology(ONTOLOGY_PATH))


def finish(run_dir: Path, rows: list[dict], meta: dict, ontology: dict) -> None:
    """Write claims.jsonl, one review_<paper>.xlsx per paper and meta.json, and print a summary."""
    reported = [row for row in rows if row["status"] == "reported"]
    meta["summary"] = {
        "claims": len(reported),
        "not_reported_relations": len(rows) - len(reported),
        "quote_not_found": sum(row.get("quote_found") is False for row in reported),
        "flagged": sum(bool(row.get("flags")) for row in reported),
        "failed_papers": sum(paper["error"] is not None for paper in meta["papers"]),
    }
    meta["finished"] = datetime.now().isoformat(timespec="seconds")
    write_jsonl(rows, run_dir / "claims.jsonl")

    # One workbook per paper, so each reviewer gets their own file.
    review_files = []

    for publication in dict.fromkeys(row["publication"] for row in rows):
        path = run_dir / f"review_{Path(publication).stem}.xlsx"
        write_review_sheet([row for row in rows if row["publication"] == publication], path, ontology,
                           meta["review"]["sample_per_paper"], meta["review"]["seed"], meta)
        review_files.append(path)

    # Papers screened by the model (no --method given) get a row in the screening sheet.
    screened = [paper for paper in meta["papers"] if paper.get("screening") is not None]

    if screened:
        path = run_dir / "screening.xlsx"
        write_screening_sheet(screened, path, meta)
        review_files.append(path)

    save_meta(run_dir, meta)

    print("\nSummary:")

    for key, value in meta["summary"].items():
        print(f"  {key}: {value}")

    print("\nReview sheets:")

    for path in review_files:
        print(f"  {path}")


def save_meta(run_dir: Path, meta: dict) -> None:
    (run_dir / "meta.json").write_text(json.dumps(meta, indent=2, ensure_ascii=False), encoding="utf-8")


def git_commit() -> str | None:
    """The current commit, with "+uncommitted" if files were changed since, so a run
    can be traced back to the exact code that produced it."""
    try:
        commit = subprocess.run(["git", "rev-parse", "--short", "HEAD"], capture_output=True, text=True, check=True)
        status = subprocess.run(["git", "status", "--porcelain"], capture_output=True, text=True, check=True)
    
    except (OSError, subprocess.CalledProcessError):
        return None
    
    return commit.stdout.strip() + ("+uncommitted" if status.stdout.strip() else "")


if __name__ == "__main__":
    main()
