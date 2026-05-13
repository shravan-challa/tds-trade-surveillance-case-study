# TD Securities — Trade Surveillance Case Study Brief

**Received:** 2026-05-11
**Deadline:** ~5 days (target completion: 2026-05-16)
**Stage:** Post first-round interview (interview went well — interviewer flagged "logical thinking" and "understanding of data")
**Tooling allowed:** ChatGPT, Claude, anything else — fully open
**Dataset:** Synthetic CSV (zipped), seeded by interviewer to mimic TD trade surveillance data

---

## Email (verbatim)

> Hi Shravan,
>
> It was very refreshing to speak with you and your understanding of data. I genuinely enjoyed your logical thinking.
>
> We do have access to ChatGPT and Claude – feel free to work on this dataset at home – it is completely synthetic in nature. There are no limitations or requirements on how you achieve your goal as long as it is defensible.
>
> Included in this email is a dataset that I seeded to mimic the data we produce within TD for surveillance.
>
> - Embedded is a defect rate to impact ability to ingest across all fields with a frequency of 1%-10% randomly selected by the program for seeding.
> - This is a Zip file that contains a CSV, the CSV has a variety of information.
>
> Priority 1 is ensuring the dataset is free of errors and has the correct format for ingestion to ensure the vendor does not drop the records which would impact completeness (JPM Case as you mentioned).
>
> Priority 2 is accuracy and efficacy of data.
>
> Priority 3 is generation of metrics.
>
> As discussed, – this is an open-ended question for priority 2, for priority 1 something are glaring some are open ended based on your imagination.
>
> Take approximately 5 days to complete.

---

## Decoded requirements

### Priority 1 — Ingestion readiness (HIGHEST WEIGHT)
- Goal: vendor surveillance system must not drop records
- Tied explicitly to the **JPM case** (which Shravan raised in interview — re: $348M Fed/OCC fine for incomplete trade surveillance data, gaps in CAT/Order Audit Trail ingestion)
- Defect rate: 1–10% per field, randomly seeded → must detect AND fix without being told what's broken
- "Glaring + open-ended" — they want both obvious mechanical fixes (nulls, type mismatches, encoding) AND judgement calls (what counts as "ingestable" vs "borderline")

### Priority 2 — Accuracy & efficacy (OPEN-ENDED)
- Explicitly flagged as "open-ended" and "based on your imagination"
- Likely: cross-field consistency, business-rule violations, statistical outliers, referential integrity, temporal coherence
- This is the differentiator section — show domain knowledge of what makes trade data *meaningful*, not just *parseable*

### Priority 3 — Metrics generation
- Surveillance KPIs: completeness %, accuracy %, timeliness, coverage by venue/product/desk
- DQ scorecard the team could actually use

---

## Strategic angles to weave in

- **JPM case is already on the table** — interviewer remembers it. Reference it explicitly in the writeup, frame the methodology as "what would have caught the JPM gap"
- **"Defensible"** is the keyword in the brief — every fix and metric needs a stated rule + rationale, not just "I cleaned this up"
- **Show the surveillance lens, not the generic DQ lens** — frame failures in terms of what they prevent regulators from seeing (e.g., spoofing, layering, wash trades, marking-the-close)
- **Reproducible pipeline** beats one-off notebook — show this is something a team could run on the next vendor file

---

## Open questions (ask interviewer if it helps)
- Is the deliverable a notebook, a report, slides, or all three? (Default: short written report + reproducible code + DQ scorecard)
- Any preferred ingestion target format? (CAT? Internal schema? Vendor-specific?)
- Should the metrics be production-style (running dashboard) or one-shot (case study output)?
