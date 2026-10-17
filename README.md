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

- `papers/dev/` — papers for development and testing
- `papers/heldout/` — papers for the final evaluation only; `run.py` refuses them unless you add `--final`

## Run

Run from the project root:

```bash
python run.py papers/dev/paper.pdf --method "Value Sensitive Design"   # method given
python run.py papers/dev/*.pdf                                       # model finds the methods
python run.py papers/heldout/*.pdf --final                           # final evaluation only
```

A paper of ~25 pages takes about 3–4 minutes (OCR is cached, so only the first run per PDF does OCR). `--method` applies to every PDF in the command. `--sample 30` changes how many quotes per paper go into the review sheet (default 20, `0` = all).

Each run creates `runs/<date_time>/`:

| File | What |
|---|---|
| `review_<paper>.xlsx` | one review workbook per paper |
| `claims.jsonl` | all claims after verification (input for the knowledge graph) |
| `extracted.jsonl` | raw model output |
| `meta.json` | ontology version, prompt hashes, model, settings, git commit, timings |

## Review (blind)

Each workbook holds a random sample of quotes the pipeline found, spread over the relations. Don't look at the Compare and Summary tabs until you have judged every quote: they show the LLM's interpretation.

1. **Judge** — for each quote, read `quote` and `context` (both straight from the paper), then pick the `human_relation` it supports (dropdown; `none` if it supports nothing) and write the `human_value` in a few words. Fill in `reviewer`. The **Relations** tab has the definitions.
2. Only when all quotes are judged, open **Compare**. It shows your decision next to the LLM's; `relation_match` fills in by itself. Set `value_match` to `yes`, `partly` or `no` for each row.
3. **Summary** shows the agreement for this paper.

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
