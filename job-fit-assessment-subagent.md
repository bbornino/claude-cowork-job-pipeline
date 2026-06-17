# Job Fit Assessment Subagent

**Purpose:** For a single shortlisted job posting, read the *actual* posting (not just the title) and compare it against `candidate-profile.md` to answer two things:

1. How close is this to a real match — what's met, what's a gap?
2. ATS keyword analysis — which terms from the posting are missing from the profile, and how critical is each one?

This subagent is invoked once per new Shortlist entry by the Daily Job Scrub Pipeline (Task 5) — it doesn't run on its own.

**Input:** One row from `job-listings.xlsx` ("Shortlist") — Job Title, Company Name, Job ID, Posting URL — plus `candidate-profile.md`.

---

## Step 1: Fetch the Full Posting

**Cowork's bash sandbox can't make outbound HTTPS calls** (confirmed: egress proxy 403s `requests` regardless of headers), so this subagent doesn't fetch anything itself either. Before invoking this subagent for a given posting, the orchestrator (Cowork, following Task 5) should:

1. Look up this company's board slug in `company-boards.xlsx` → "Board URL" (e.g. `job-boards.greenhouse.io/stripe` → slug `stripe`).
2. `web_fetch` `https://boards-api.greenhouse.io/v1/boards/{slug}/jobs/{Job ID}?content=true`.
3. Run `python3 daily_pipeline.py extract-cache posting {Job ID} <result_file>`, which writes `{cache_dir}/postings/{Job ID}.json` (default cache dir `/tmp/pipeline_cache`).

This subagent then reads that cache file. Three possible outcomes:

- **Normal:** the cache file is a full job-detail object with a `content` field (HTML job description) plus `location`, `absolute_url`, etc. Extract from `content`:
  - Required qualifications (however labeled — "Requirements," "Must have," "What you'll need," etc.)
  - Preferred/nice-to-have qualifications (separately from required — these carry less weight)
  - Any explicit list of tools/technologies/skills mentioned anywhere in the posting, not just in a bulleted requirements section — job descriptions often mention stack details in the role description itself

- **`{"_status": "empty"}`** — the Greenhouse API returned no content at all for this Job ID (seen for ~2/8 postings in one run; likely the posting was closed/drafted/paused between when it was shortlisted and now). **Don't retry more than once** — if a second `web_fetch` + `extract-cache` also produces `"empty"`, skip the requirements/ATS analysis entirely and go straight to Step 5 with: `Requirements Met (Required) = N/A`, `Requirements Met (Preferred) = N/A`, `Key Gaps = N/A`, `ATS Keywords - Present/Missing (Critical)/Missing (Optional) = N/A`, `Overall Fit Note = "Posting unavailable - API returned empty response. Verify at posting URL or skip."` (`Pay Flag`/`Pay Range` are unaffected — Step 2 reads those from the Shortlist row, not from this fetch, so run Step 2 regardless of this outcome.)

- **`{"_status": "truncated"}`** — the response was cut off mid-object before `extract-cache` could parse even the fields before `content` (rare - `content` is usually the last/largest field, so partial description text is normally still usable; this only happens if the cut lands earlier). Treat the same as `"empty"` above, but with `Overall Fit Note = "Posting unavailable - response truncated before any usable content. Verify at posting URL or skip."`

Fall back to `web_fetch` directly on the Posting URL only if Job ID is blank or the company isn't a Greenhouse source — in which case the same egress restriction may apply; if so, treat as the empty-response case above.

---

## Step 2: Pay Range Check — Do This First

Read the Shortlist row's `Pay Range` field (already captured in Task 1 — no extra fetch needed, this doesn't depend on Step 1 at all and can be checked before or in parallel with it).

- **`Not Disclosed`** → `Pay Flag = Not Disclosed`. Continue to Step 3 as normal.
- Otherwise, estimate an annual range using the conversion rule in `settings.md` → "Compensation Thresholds" (hourly × 2,080 for annual equivalent). Use the current floor/target/unusually-high values from that file (as of this writing: $90k / $120k / $180k):
  - Annual **max** < floor ($90,000) → `Pay Flag = Too Low`
  - floor ≤ Annual **max** < target ($120,000) → `Pay Flag = Below Target`
  - Annual **min** ≥ unusually-high threshold ($180,000) → `Pay Flag = Unusually High`
  - Otherwise → `Pay Flag = OK`

**If `Pay Flag` is `Too Low` or `Unusually High`, stop here — skip Steps 3 and 4 entirely** and go straight to Step 5 with `Requirements Met (Required) = N/A`, `Requirements Met (Preferred) = N/A`, `Key Gaps = N/A`, `ATS Keywords - Present/Missing (Critical)/Missing (Optional) = N/A`, and an `Overall Fit Note` along the lines of "Pay ($X-Y) is below the $90k floor — not pursuing" or "Pay ($X-Y) is unusually high for this title — likely scoped more senior than posted, or an inflated/placeholder range; not worth the analysis time." This is the main token-saving move in this subagent: pay is the cheapest check (no fetch needed) and the most decisive filter, so it runs before the expensive requirements/ATS work rather than after it.

For `OK`, `Below Target`, or `Not Disclosed`, continue to Step 3.

---

## Step 3: Requirements Match

For each **required** item, classify against `candidate-profile.md`:
- `MEET` — directly covered (production experience or clearly equivalent)
- `PARTIAL` — related/transferable experience, but not a direct match (e.g. PHP background for a "Python required" posting)
- `GAP` — no real basis for this in the profile

Do the same for **preferred** items, but keep them in a separate tally — a gap in "preferred" matters much less than a gap in "required."

Output:
- `Requirements Met (Required): X/Y`
- `Requirements Met (Preferred): X/Y`
- `Key Gaps:` short list of the GAP and notable PARTIAL items, one line each

---

## Step 4: ATS Keyword Analysis

**Skip this step if Pay Flag is `Too Low` or `Unusually High`** (per Step 2) — go straight to Step 5.

Pull the specific skill/tool/technology terms and key phrases from the posting (Step 1) — these are the kinds of terms an ATS scan matches on. Compare each against `candidate-profile.md`'s ATS Keyword Bank and overall content.

For each posting keyword:
- **Present** — already covered in the profile, no action needed
- **Missing — Critical** — appears in the *required* section, or appears multiple times across the posting (signals it matters to this employer specifically)
- **Missing — Optional** — appears only in "preferred"/nice-to-have, or mentioned once in passing

Only flag something as missing if it's a term the candidate could honestly claim or quickly pick up — don't flag things that would be an overstatement (cross-reference "Notes for Tailoring" in `candidate-profile.md`).

Output:
- `ATS Keywords - Present:` comma-separated
- `ATS Keywords - Missing (Critical):` comma-separated
- `ATS Keywords - Missing (Optional):` comma-separated

---

## Step 5: Write Result

Add one row to `job-listings.xlsx` → "Fit Assessment":

| Column | Source |
|---|---|
| Job ID | input (join key back to All Listings - preferred match) |
| Posting URL | input (fallback join key if Job ID is blank) |
| Company Name | input |
| Job Title | input |
| Requirements Met (Required) | Step 3, or `N/A` if skipped per Step 2 |
| Requirements Met (Preferred) | Step 3, or `N/A` if skipped |
| Key Gaps | Step 3, or `N/A` if skipped |
| Pay Range | the Shortlist row's `Pay Range` field, carried through as-is — this is the actual number; `Pay Flag` (below) is the verdict against it |
| Pay Flag | Step 2 — `OK` / `Below Target` / `Too Low` / `Unusually High` / `Not Disclosed` |
| ATS Keywords - Present | Step 4, or `N/A` if skipped |
| ATS Keywords - Missing (Critical) | Step 4, or `N/A` if skipped |
| ATS Keywords - Missing (Optional) | Step 4, or `N/A` if skipped |
| Overall Fit Note | one line. If pay-filtered (`Too Low`/`Unusually High`): just the pay reason, e.g. "Pay ($60-70k) is Too Low — below the $90k floor; not pursuing." Otherwise: incorporate the Pay Flag if not `OK`/`Not Disclosed`, e.g. "Strong match on skills, but pay ($95-110k) is Below Target" |
| Date Assessed | today's date |

**Also write the job description to `All Listings`.** Step 1 already ran `extract-cache posting {Job ID} <result_file>`, which saved the full posting (including the `content` field) to the local cache. Just run:

```
python3 daily_pipeline.py write-description {Job ID}
```

This reads the cached posting, strips the HTML to plain text (capped at 2,000 chars), and writes it to the `Job Description` column in `All Listings` (matched by Job ID) — no need to paste the HTML itself anywhere. Run this immediately after writing the Fit Assessment row, before moving to the next posting — don't batch them at the end. If the cached posting was empty/unavailable (Step 1's outcome 2 or 3), this command reports that and writes nothing — skip it for that posting.
