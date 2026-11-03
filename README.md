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
python run.py papers/dev/*.pdf --methods-from runs/<date_time>          # same methods as an earlier run, no screening
```

`--methods-from` reuses the methods an earlier run included per paper (matched on file content, so renamed PDFs still match). It was used for round 2, so that both rounds extract the same methods and only the revised definitions and instructions differ.

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

Each workbook holds a random sample of quotes the pipeline found, spread over the relations. Don't look at the Compare and Summary tabs until you have judged every quote: they show the LLM's interpretation. Only the cells you fill in are editable; everything else (quotes, the LLM's output and the grey formula cells) is locked, because typing over a formula stops the sheet from updating.

1. **Judge** — for each quote, read `quote` and `context` (both straight from the paper). Then put an `x` in the column of **every** relation the quote supports for the method in the `method` column — as many as apply. If it says nothing about that method (e.g. when it is about another method in the same paper), put the `x` in `supports_nothing` instead, so the row counts as judged. Judge **the quote, using the context only to understand it**: mark what the quote itself says, and use the context to see what it refers to and whether it is about the method. Something that only the surrounding sentences say does not count for this quote. Motivation sentences (about a problem or a gap in the field) only count if the context links them to the method. (Example: "New technologies often give rise to previously unknown ethical problems…" says nothing about eTA on its own, but the next sentence proposes eTA as the answer to it, so it supports `addresses`.) Write the main value in `human_value` in a few words; `notes` and `reviewer` are optional. The **Relations** tab has the definitions.
2. Only when all quotes are judged, open **Compare**. It shows your relations next to the LLM's: `agreed`, `llm_only` (the LLM gave it, you didn't) and `human_only` (you gave it, the LLM didn't) fill in by themselves. For rows with an agreed relation, set `value_ok` to `yes`, `partly` or `no`: are the LLM's values for those relations right? And where `llm_only` lists relations, look again: if the LLM saw something you missed, type in `llm_only_correct` **how many** of those relations are right after all (leave it empty for 0). This gives a second set of "corrected" numbers next to the blind ones; your own marks on Judge stay as they are.
3. **Summary** shows the agreement for this paper.

**Screening** (`screening.xlsx`, same idea): the **Judge** tab has one row per method the LLM found (or `(no method found)`). Per method, judge each criterion **C1–C3** as `met`, `not met` or `not in this paper` (the paper gives no evidence either way; for C3, evidence of an ICT application may be in another publication), then decide:

- `include`: all criteria met;
- `exclude`: a criterion is not met, or C1 or C2 cannot be judged from this paper (it does not describe the method well enough to extract it);
- `pending`: C1 and C2 are met, but the evidence for C3 has to come from another publication. The method is then decided across all papers about it, or by the team.

`evaluate.py` derives this reference decision from the C1–C3 judgements by these rules (`reference_decision`), so it is applied the same way to every method; the typed decision is only used for rows without criteria, such as `(no method found)`. Those rows are reported per paper, not counted as methods.

If the paper describes a method that isn't listed, add a row at the bottom with the paper, the method and your decision; with `include` it counts as missed by the LLM. Then open **Compare** for the LLM's decision and its judgement and reason per criterion. Rebuilding a run keeps the screening judgements.

The screening check only covers methods the LLM found, plus any a reviewer happens to add: it does not measure how many methods the LLM missed.

When all papers of a run are reviewed, add them up:

```bash
python evaluate.py runs/<date_time> --complete-only                    # one run
python evaluate.py runs/<round1> runs/<round2> --complete-only         # compare two rounds
python evaluate.py ... --weighted                                      # also weighted estimates (see below)
```

Besides the printed scores, it writes three files to each run folder: `results.tex` (the numbers as LaTeX commands for the paper), `review_overview.md` (raw counts per paper and per relation, the screening decisions and all reviewer notes, for finding error patterns) and `scores.json`. In `results.tex`, every measure is a percentage plus its counts: `\ResPrecision{} (\ResPrecisionN{})` gives "54% (89/166)". With two runs the macros are `\ResOne...` and `\ResTwo...`, so the paper can `\input` both files, and the second run's file also holds `\ResComparisonTable`. `--complete-only` leaves out papers whose review is not finished. It reports these measures, each in its own unit:

- **exactly the same relations**, per quote;
- **relation precision**, per (quote, relation) pair the LLM gave: how many you also gave (catches over-extraction);
- **relation recall**, per (quote, relation) pair you gave: how many the LLM also gave. This is recall *within the quotes the pipeline extracted*, not completeness across the paper: passages the pipeline never extracted are never reviewed;
- **values right**, per quote with at least one agreed relation.

Pairs are not claims: one pair can hold several claims (three activities named in one quote), so none of the measures is "per claim".

It also reports how many quotes **support nothing** for the method, and how many **LLM-only relations** the reviewer accepted after comparison. Exact agreement, precision and recall are also reported **corrected** with `llm_only_correct`: relations the LLM found and the reviewer first missed then count as right. The blind numbers are the strict ones; the corrected numbers account for reviewer misses, but were decided after seeing the LLM's answer, so report both.

The numbers **describe the reviewed sample** (sums over the reviewed quotes, with their counts). The sample is stratified: spread over (method, relation) groups, so rare relations are over-represented and the sample is not proportional to all extracted quotes. With `--weighted`, `evaluate.py` additionally estimates the results for all extracted quotes: each sampled quote is weighted by the number of extracted quotes it stands for (inverse inclusion probability, reconstructed from `claims.jsonl` and the run's seed; groups with no sampled quote are covered by scaling the paper's weights to its total), with 95% intervals from a bootstrap that resamples quotes within each paper. These estimates (`\ResWeighted...`) were not used for the paper, because the validity of the weighting for this design was not confirmed.

## Changing things

- **Ontology** (concepts, definitions, examples, relations): edit `config/ontology.yaml` and raise its `version`. Prompt and schema update automatically. `enabled: false` switches a concept off.
- **Prompt wording**: `prompts/extract.md` and `prompts/identify.md`.
- **Model and settings**: constants at the top of `extract.py` (extraction) and `ocr.py` (OCR).
- After changing `verify.py` or `export.py`, redo an old run for free: `python run.py --rebuild runs/<date_time>`. This overwrites that run's review workbooks (same sample), so copy them first if someone already started reviewing.

Commit before a run you want to keep: `meta.json` records the git commit and marks it `+uncommitted` otherwise.
