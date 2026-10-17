# Method extraction tool

Extracts ontology-aligned claims about ethical and sustainable ICT reasoning methods from PDFs (Mistral OCR + Mistral Large), checks every quote against the paper, and exports an Excel sheet for human review.

## Setup

```bash
python -m venv .venv
source .venv/bin/activate          # fish: source .venv/bin/activate.fish
pip install -r requirements.txt
cp .env.example .env               # then put your MISTRAL_API_KEY in .env
```

## Papers

Drop the PDFs in one of two folders:

- `papers/dev/` — papers you may learn from: review their errors and improve the prompt or ontology based on them
- `papers/heldout/` — papers kept out of development, only used to measure the final result; `run.py` refuses them unless you add `--final`

## Run

Run from the project root:

```bash
python run.py                                                        # every PDF in papers/dev/
python run.py --final                                                # every PDF in papers/heldout/
python run.py papers/dev/paper.pdf --method "Value Sensitive Design"   # one paper, method given
```

Without `--method`, the model first **screens** each paper: it lists the methods it describes and checks each against the inclusion criteria (C1–C3 in `config/ontology.yaml`). Only methods that meet all criteria are extracted; a paper without one is rejected. With `--method`, screening is skipped.

A paper of ~25 pages takes about 3–6 minutes (OCR is cached, so only the first run per PDF does OCR). Each run processes all papers again: a run is one complete, traceable snapshot. Temporary API errors are retried automatically. If a paper still fails (`failed_papers` in the summary), redo just those papers inside the same run with `python run.py --retry-failed runs/<date_time>`; this only works if the ontology, prompts and settings haven't changed since. `--method` applies to every PDF in the command. `--sample 30` changes how many quotes per paper go into the review sheet (default 20, `0` = all).

Each run creates `runs/<date_time>/`:

| File | What |
|---|---|
| `review_<paper>.xlsx` | one review workbook per included paper |
| `screening.xlsx` | one row per screened paper: include or exclude |
| `claims.jsonl` | all claims after verification (input for the knowledge graph) |
| `extracted.jsonl` | raw model output |
| `meta.json` | ontology version, prompt hashes, model, settings, git commit, timings |

## Review (blind)

Each workbook holds a random sample of quotes the pipeline found, spread over the relations. Don't look at the Compare and Summary tabs until you have judged every quote: they show the LLM's interpretation.

1. **Judge** — for each quote, read `quote` and `context` (both straight from the paper), then pick the `human_relation` it supports for the method in the `method` column (dropdown; `none` if it supports nothing about that method, e.g. when it is about another method in the same paper) and write the `human_value` in a few words. Fill in `reviewer`. The **Relations** tab has the definitions.
2. Only when all quotes are judged, open **Compare**. It shows your decision next to the LLM's; `relation_match` fills in by itself. Set `value_match` to `yes`, `partly` or `no` for each row.
3. **Summary** shows the agreement for this paper.

**Screening** (`screening.xlsx`, same idea): the **Judge** tab has one row per method the LLM found (or `(no method found)`). Decide per method `include` or `exclude`: does it meet all criteria? If the paper describes a method that isn't listed, add a row at the bottom with the paper, the method and `include`; it counts as missed by the LLM. Then open **Compare** for the LLM's decision and its reason per criterion.

When all papers of a run are reviewed, add them up (with 95% confidence intervals):

```bash
python evaluate.py runs/<date_time>                    # one run
python evaluate.py runs/<baseline> runs/<improved>     # compare two runs
```

## Changing things

- **Ontology** (concepts, definitions, examples, relations): edit `config/ontology.yaml` and raise its `version`. Prompt and schema update automatically. `enabled: false` switches a concept off.
- **Prompt wording**: `prompts/extract.md` and `prompts/identify.md`.
- **Model and settings**: constants at the top of `extract.py` (extraction) and `ocr.py` (OCR).
- After changing `verify.py` or `export.py`, redo an old run for free: `python run.py --rebuild runs/<date_time>`. This overwrites that run's review workbooks (same sample), so copy them first if someone already started reviewing.

Commit before a run you want to keep: `meta.json` records the git commit and marks it `+uncommitted` otherwise.
