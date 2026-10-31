# Evaluation of the extraction: sampling design and estimators

Short note for checking the statistical approach. Implementation: `evaluate.py`
(`sampling_weights`, `design_estimates`); sampling: `export.py` (`review_items`, `sample_items`).

## Population and units

- **Population**: all *extracted quotes* of the reviewed publications. A quote is a verbatim passage
  that the pipeline linked to one or more claims for one method. Claims with the same method and the
  same quote form one quote.
- A **claim** is (method, relation, value, quote). One quote can hold several claims, also for the
  same relation (e.g. three activities named in one sentence).
- The review judges, per quote, the **set of relations** the quote supports. Measures are therefore
  defined on quotes and on **(quote, relation) pairs**, not on claims.

Round 1: 6 publications, 307 extracted quotes in total (41 to 64 per publication), 120 reviewed.

## Sampling design

Per publication, 20 quotes (all if fewer):

1. Quotes are grouped into strata by (method, relation of the quote's first claim).
2. Strata are visited in a random order, repeatedly; each visit takes one quote at random from the
   stratum, until 20 quotes are taken. Strata are reshuffled before each pass. A fixed seed makes the
   sample reproducible.

This gives each stratum about the same number of quotes, so rare relations are over-represented
relative to the population. When a publication has more than 20 strata, some strata get no quote in
the sample (round 1: 1 of 49 quotes in one publication, 23 of 64 in another).

## Weights

For a sampled quote in stratum *g* of publication *p*, with *N_g* quotes in the stratum and *n_g*
of them sampled:

    w_raw = N_g / n_g                                  (inverse inclusion probability, given the allocation)
    w     = w_raw × N_p / Σ_sampled-quotes-of-p w_raw  (N_p = quotes in publication p)

The second step spreads the quotes of strata without a sampled quote over the publication's
sampled quotes; it assumes those strata resemble the rest of the publication.

## Estimators

Every measure is a ratio estimator over all reviewed publications together:

    estimate = Σ_q w_q · y_q  /  Σ_q w_q · x_q

| Measure | y_q | x_q | Unit |
|---|---|---|---|
| Exact agreement | 1 if reviewer's relation set = LLM's relation set | 1 | quote |
| Relation precision | \|H_q ∩ L_q\| | \|L_q\| | (quote, relation) pair given by the LLM |
| Relation recall | \|H_q ∩ L_q\| | \|H_q\| | (quote, relation) pair given by the reviewer |
| Value correctness | 1 if the LLM's values for the agreed relations are judged correct | 1 if \|H_q ∩ L_q\| > 0 | quote with an agreed relation |

*H_q* and *L_q* are the reviewer's and the LLM's relation sets for quote *q*. Corrected variants add
the number of LLM-only relations the reviewer accepted after seeing the LLM's answer to the
numerator (and, for recall, to the denominator).

Recall is recall **within extracted quotes**: passages the pipeline did not extract are not in the
population, so it does not measure completeness across publications.

## Confidence intervals

95% percentile bootstrap, 2000 replicates: within each publication, the reviewed quotes are
resampled with replacement (keeping their weights), and the estimator is recomputed over all
publications. Resampling whole quotes keeps the relations of one quote together, so the dependence
between pairs of the same quote is respected.

## Questions to check

1. Is the inverse-probability weight *N_g / n_g* appropriate, given that the allocation over strata
   (which strata get one quote more) is itself random?
2. Is scaling within a publication an acceptable way to handle strata without a sampled quote, or
   should strata be collapsed beforehand (e.g. per relation instead of per method and relation)?
3. Does resampling within publications (keeping weights fixed) give adequate intervals for these
   ratio estimators, with 20 quotes per publication and some strata with a single quote?
4. Should reviewer disagreement (single reviewer per quote, corrected by a second) be reflected in
   the intervals, or reported separately?
