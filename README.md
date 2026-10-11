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

A paper of ~25 pages takes about 3–4 minutes (OCR is cached, so only the first run per PDF does OCR). `--method` applies to every PDF in the command.

Each run creates `runs/<date_time>/`:

| File | What |
|---|---|
| `review.xlsx` | the sheet to review |
| `claims.jsonl` | all claims after verification (input for the knowledge graph) |
| `extracted.jsonl` | raw model output |
| `meta.json` | ontology version, prompt hashes, model, settings, git commit, timings |

## Review

Open `review.xlsx`:

- **Claims** — one row per claim. Red = quote not found in the paper (likely hallucinated), yellow = other flag (see `flags`). Read `quote` and `context`, then fill in `verdict` (dropdown), `correction`, `notes` and `reviewer`.
- **Not reported** — relations the model says the paper doesn't cover. Mark each as `correct (not in paper)` or `missed (in paper)`.
- **Run info** — the settings of this run.

## Changing things

- **Ontology** (concepts, definitions, examples, relations): edit `config/ontology.yaml` and raise its `version`. Prompt and schema update automatically. `enabled: false` switches a concept off.
- **Prompt wording**: `prompts/extract.md` and `prompts/identify.md`.
- **Model and settings**: constants at the top of `extract.py` (extraction) and `ocr.py` (OCR).
- After changing `verify.py` or `export.py`, redo an old run for free: `python run.py --rebuild runs/<date_time>`.

Commit before a run you want to keep: `meta.json` records the git commit and marks it `+uncommitted` otherwise.
