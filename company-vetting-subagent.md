# Company Vetting Subagent

**Purpose:** Given a single company name that `daily-job-scrub-pipeline.md` → Task 3 has already triaged as needing fresh vetting (not blacklisted, and either new, `PENDING`, or past its `Next Review Date`), run a full evaluation (Steps 1–4) and write/update its row in `company-tracker.xlsx`.

This subagent is invoked only for companies Task 3's triage selected, within the per-run cap. It is **not** invoked for companies that resolve instantly during triage (already `BLACKLIST`, or `PURSUE`/`WATCH` within their review window) — those never reach this subagent at all, so there's no tracker check to repeat here. Go straight to Step 1.

---

## Step 1: Company Research (Source of Truth)

Gather raw facts once. Steps 2–4 interpret this data — no re-researching the same things twice.

**Search budget:** aim for roughly 1–2 targeted searches per snapshot area below. If a fact isn't surfaced after that, mark `[UNKNOWN]` and move on — `[UNKNOWN]` is an expected, valid result (see "Important" note below), not a failure. Don't keep searching for something that may simply not be public.

**Consolidate sources.** Many of these areas can come from the same fetch — a company's official "About" or "Careers" page often covers size, HQ, and remote policy in one fetch; a single search like "[Company] layoffs funding 2025 2026" often covers hiring trend and ghost-job signals together. Aim for roughly 2-4 total fetches/searches per company, not one per field.

**[COMPANY SNAPSHOT]**
- **Size:** employee count (note source — LinkedIn counts are often stale/approximate)
  - Sweet spot: 150–500
  - Flag if <100 (high chaos risk) or >1500 (bureaucracy risk)
  - Note 500–1500 as a yellow flag
- **HQ + offices:** city/state for all major locations — this is a hard-gate input (see Step 2), not background color, so get this right even if other fields come back thin
- **Remote policy:** current stated policy AND most recent news on RTO mandates — check these separately, since policies have shifted a lot in 2024–2025
- **Funding/runway signal only** (not the specific round/amount/date — that level of detail doesn't change any verdict): is there any indication of funding trouble or instability? A clean "no concerns found" is a complete answer here.
- **Hiring activity:** number of currently open roles, and any hiring sprees / freezes / layoffs in the last 12 months (cite source + date)
- **LGBTQ+/DEI signals:** HRC Corporate Equality Index score if available, stated gender-affirming care coverage, notable employee review mentions
- **AI/engineering culture:** do job postings or the eng blog mention AI tooling (Copilot/Cursor/Claude/agentic workflows)?
- **Application limit:** does this company restrict how many roles a candidate can apply to at once or within a time window (e.g. "one active application at a time," "max 3 roles per 6 months")? Check careers FAQ pages and the boilerplate text on job postings themselves — this is often stated there. If found, record the number (and the window, if any) in "Source / Notes." If nothing found, mark `[UNKNOWN]` — don't assume unlimited.
- **Prior application history:** check `application-history.xlsx` for this company name. This is a B-maintained record of "I've applied here before and here's what happened" — no web research needed, just a lookup. If a row exists, note `Times Rejected` for use in Step 2.

**Important:** if a field can't be found, mark it `[UNKNOWN]` explicitly. Do not leave it blank, and do not assume neutral or positive by default.

---

## Step 2: Fit & Safety Assessment

Use only the Step 1 snapshot — no new research.

**Hard No — stop here if any apply:**
- Company is Meta, Tesla, SpaceX, or X
- Clear evidence of a trans-hostile environment (poor/no CEI score + no gender-affirming care coverage + negative LGBTQ+ employee reviews, or news of DEI/benefits rollbacks targeting LGBTQ+ employees)
- **HQ is in a state currently rated "Do Not Travel" on Erin Reed's (Erin in the Morning) anti-trans legal risk map** — as of Feb 2026: **Florida, Kansas, Texas**. This list changes; if it's been a while since this was last checked against erininthemorning.com, verify the current list before relying on it rather than assuming these three are still accurate. This is a legal-jurisdiction risk, not a remote-culture judgment — a company with an otherwise-great remote setup still gets excluded if its HQ state is in this category, since a harassment/discrimination claim would likely be litigated there.
- Security clearance is required for the role(s) and I don't hold one
- No remote option AND offices are outside commute range of Sacramento, CA

**If no Hard No applies, assess:**
- **Remote fit:** stated policy + RTO trend — clean, workable, or concerning?
- **LGBTQ+/trans safety:** based on CEI score, stated benefits, and employee reviews — not HQ location alone for THIS signal specifically (a company HQ'd in a hostile state can still have a genuinely safe remote culture, and vice versa).
- **Legal jurisdiction risk (HQ Location), below the Hard No threshold:** the Do Not Travel states above are an outright exclusion, not a downgrade. For HQ states one tier down on Erin Reed's map (currently labeled something like "Worst States" — meaningful anti-trans legislation but not yet a Do Not Travel rating), apply a downgrade rather than exclusion: if a harassment/discrimination claim ever became necessary, litigation would likely happen in the company's HQ state, so this remains a real factor even short of the hard cutoff. Downgrade Company Fit by one tier for HQ in this category. When choosing between two otherwise-comparable remote opportunities and only one carries this downgrade, prefer the other, even at some cost to pay or other factors. Note this explicitly in Reason/Summary (e.g. "Downgraded to Solid Contender — HQ in [state] carries elevated litigation-jurisdiction risk despite a decent remote culture").
- **Pedigree risk:** concrete evidence only — e.g. explicit "top university" language, eng team composition skewing heavily FAANG. If there's no concrete evidence, mark `[NO SIGNAL]` rather than assuming risk from vibes.
- **AI/tooling culture fit:** does their engineering culture align with an AI-adjacent positioning?
- **Size fit:** per the Step 1 thresholds
- **Prior rejection history:** if `application-history.xlsx` shows `Times Rejected` for this company —
  - **2:** downgrade Company Fit by one tier (Strong Match → Solid Contender, Solid Contender → Long Shot). Floor at Long Shot — don't go to Hard No for this reason alone.
  - **3+ ("a bunch"):** downgrade by two tiers, same floor at Long Shot.
  - Note the downgrade explicitly in the Reason/Summary, e.g. "Downgraded from Strong Match to Long Shot — 3+ prior ATS-stage rejections." This isn't permanent: a substantially different role type (e.g. backend-focused vs. the frontend roles previously applied to) could warrant revisiting at the next `Next Review Date`.

**[FIT: Strong Match / Solid Contender / Long Shot / Hard No]**

---

## Step 3: Ghost Job / Hiring Health Check

Use the Step 1 snapshot. For each item below, output `[CLEAR] / [FLAG] / [RED FLAG] / [UNKNOWN]`.

`[UNKNOWN]` is a valid and expected result — do not default to `[CLEAR]` just because nothing turned up.

- **Posting volume:** many simultaneous roles → possible pipeline-building vs. real need
- **Layoffs / freeze / funding trouble:** in the last 6–12 months
- **Funding runway:** any concerns
- **External chatter:** Glassdoor/Blind/LinkedIn intel suggesting this company's hiring isn't real
- **Recurring-posting pattern:** evidence this company re-posts roles repeatedly over many months — a company-wide tendency, not tied to one specific posting

**[GHOST VERDICT: LIKELY REAL / UNCERTAIN / LIKELY GHOST]**
- LIKELY REAL requires at least one piece of *positive* evidence (active growth, recent funding, a genuine hiring need) — not just an absence of red flags
- Default to UNCERTAIN if most checks come back UNKNOWN

---

## Step 4: Final Verdict & Tracker Update

Write/update one row in `Company Tracker` with these columns:

| Column | Source |
|---|---|
| Company Name | input |
| Status | derived — see logic below |
| Company Fit | Step 2 |
| Ghost Verdict | Step 3 |
| Reason / Summary | one line, written here |
| Date Checked | today's date |
| Review Interval (months) | default 6 — set higher (e.g. 12) for companies where fit/safety signals are unlikely to change quickly |
| Next Review Date | `=EDATE(Date Checked, Review Interval)` — auto-calculated by the spreadsheet |
| HQ Location | Step 1 — drives the Hard No / downgrade logic in Step 2, get this one right |
| Employee Count | Step 1 |
| Size Flag | Step 1 |
| Remote Policy & RTO Trend | Step 1 |
| Hiring Trend (12mo) | Step 1 — fold any funding-runway concern in here rather than a separate funding field, since the runway signal (not the round size/date) is what actually matters |
| LGBTQ+/DEI Signal | Step 1 |
| AI/Eng Culture Signal | Step 1 |
| Pedigree Risk | Step 2 |
| Prior Rejections | Step 1 — `Times Rejected` from `application-history.xlsx`, blank if no entry |
| Application Limit | Step 1 — a number if known (e.g. `1`), blank/`[UNKNOWN]` if not found |
| Source / Notes | application-limit window details, plus any citation that materially supports a Hard No / downgrade decision — not a general link dump |

**Status logic:**
- **PURSUE:** Strong Match or Solid Contender, AND Ghost Verdict is LIKELY REAL or (UNCERTAIN with no red flags)
- **WATCH:** Solid Contender/Long Shot with unresolved UNKNOWNs — revisit if a specific posting looks unusually compelling
- **BLACKLIST:** Hard No triggered, OR (Long Shot + LIKELY GHOST), OR multiple RED FLAGs

If the company was already in the tracker, update the existing row rather than adding a duplicate.
