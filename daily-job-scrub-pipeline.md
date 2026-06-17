# Daily Job Scrub Pipeline

Runs once per day. Takes a large raw batch of new job postings and narrows it down to a short list worth actually looking at — by scrubbing both the **job titles** and the **companies** behind those postings.

**Seven files, each doing one job:**
- `settings.md` — tunable numbers (volume caps, compensation thresholds) and the color-coding spec. Check this first.
- `company-boards.xlsx` — one row per company's job board: which company, which platform, which URL. This is *what Task 1 actually pulls from* — "pull from the companies listed in the spreadsheet."
- `job-sources.xlsx` — list of platforms (Greenhouse, Lever, etc.) with a Status (Enabled/Considering/Avoid). A company board in `company-boards.xlsx` only gets pulled if its Source is `Enabled` here — this is the "slowly turn on more sources" dial.
- `job-title-filters.xlsx` — keyword rules for PURSUE/SKIP/CHECK by title. Rarely changes; B-maintained.
- `location-filters.xlsx` — Sacramento Area / SF Bay Area city keywords. Rarely changes; B-maintained.
- `seen-postings.xlsx` — every individual posting ever pulled, by URL + Source ("Seen Postings"), plus a "Needs Link" sheet for postings that couldn't get a URL.
- `company-tracker.xlsx` — one row per company, the durable company-level verdict (see `company-vetting-subagent.md`). Grows slowly.
- `job-listings.xlsx` ("Shortlist" + "Fit Assessment") — daily output, plus per-posting fit details once assessed.

---

## Running the Pipeline: `daily_pipeline.py`

Tasks 0, 1, 2, Task 3's triage, and Task 4 are implemented once, as a script (`daily_pipeline.py`), rather than re-derived by Cowork each run. They're all deterministic data operations — dedup by Job ID/URL, keyword/substring matching against `job-title-filters.xlsx` and `location-filters.xlsx`, lookups against `company-tracker.xlsx`, counts, and spreadsheet reads/writes. The script reads its tunable values from the YAML config block at the top of `settings.md`.

**Confirmed:** Cowork's bash sandbox cannot make outbound HTTPS calls (egress proxy returns 403 for `requests`, regardless of headers — not a Greenhouse-side block). So the script never fetches anything itself by default. Cowork pre-fetches everything via its own `web_fetch`, normalizes the result with `extract-cache`, and the script reads from that cache. This also solves `web_fetch`'s ~94KB truncation on large boards — `extract-cache` extracts every *complete* job object before the cut point; any trailing jobs just get picked up on a later run via dedup, nothing is lost.

**A daily run is five steps:**

1. **Pre-fetch every board.** For each row in `company-boards.xlsx` where `Status = Active` and its `Source` is `Enabled` in `job-sources.xlsx`: get the slug (`slug_from_board_url`, works on either `job-boards.greenhouse.io/{slug}` or the API form), `web_fetch` `https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=false`, save the result (or, if `web_fetch` errors due to size, its saved tool-result file) to a local file, then run:
   ```
   python3 daily_pipeline.py extract-cache board {slug} <result_file>
   ```
   This writes `{cache_dir}/boards/{slug}.json`. Default cache dir is `/tmp/pipeline_cache` (Linux ephemeral filesystem — avoids a Windows-mount issue where files created/copied this session can't be overwritten or deleted from the bash sandbox). Override with `--cache-dir` or the `PIPELINE_CACHE_DIR` env var if needed, but keep it off any Windows-mounted path.

2. **`python3 daily_pipeline.py phase1`** — Tasks 0, 1, 2, Task 3 triage. Reads each board's cache file, dedups, applies the location and title filters, and triages every company against `company-tracker.xlsx`. Writes `seen-postings.xlsx`, `job-title-filters.xlsx` (Unclassified Titles), and `company-tracker.xlsx` (PENDING rows for over-cap companies). Outputs `survivors.json` (this run's surviving postings, for step 4) and `handoff_companies.json` (companies needing full vetting this run, capped at `max_new_companies_per_run`). A board with no cache file is reported as a "board error" — not fatal, just means that board contributes 0 postings this run.

   **Checkpoint.** `seen-postings.xlsx`, `job-title-filters.xlsx`, `company-tracker.xlsx`, `survivors.json`, `handoff_companies.json`, and `phase1_report.json` now hold everything needed to continue. Nothing about *how* each board's cache was fetched/extracted, or the per-posting filtering detail, is needed again. **If Cowork's context can be compacted/summarized at a chosen point, this is one** — the only things that matter going forward are the printed Task 0-3 summary line and the contents of `handoff_companies.json`.

3. **Run `company-vetting-subagent.md`** for each company in `handoff_companies.json`, writing each verdict as a row in `company-tracker.xlsx` (new companies get a new row; this is the only manual/LLM step in phase 1). If `handoff_companies.json` is empty, skip straight to step 4.

   **Checkpoint.** Every verdict is now a row in `company-tracker.xlsx`. The research that produced each verdict (search results, pages fetched, reasoning) doesn't need to be retained — `company-tracker.xlsx` is the durable record. **Another good compaction point**, especially since this step can involve the most varied/heaviest web research of the whole run.

4. **`python3 daily_pipeline.py phase2`** — Task 4. Re-reads `survivors.json` and the now-current `company-tracker.xlsx`, filters by company Status/Company Fit/Application Limit, and adds qualifying postings to `job-listings.xlsx` ("All Listings"). Also rescans `seen-postings.xlsx` for any "company pending - revisit" postings whose company resolved out of PENDING this run — for those, pre-fetch the posting detail the same way (`web_fetch` `.../jobs/{job_id}?content=true`, then `python3 daily_pipeline.py extract-cache posting {job_id} <result_file>`) *before* running `phase2`, if you want that rescan to fully resolve rather than fall back to "check manually." This is rare; usually `handoff_companies.json` (and therefore this rescan) is empty or small. Outputs `handoff_fit_assessments.json` — Shortlist rows needing `job-fit-assessment-subagent.md`, prioritized by Company Fit then Role Match, capped at `max_fit_assessments_per_run`.

   **Checkpoint.** `job-listings.xlsx` ("All Listings"), `seen-postings.xlsx`, `handoff_fit_assessments.json`, and `phase2_report.json` now hold everything needed to continue. **Third compaction point** — the only thing that matters going forward is the contents of `handoff_fit_assessments.json`.

5. Run `job-fit-assessment-subagent.md` for each row in `handoff_fit_assessments.json`, writing results to `job-listings.xlsx` ("Fit Assessment"). This subagent also needs each posting's content — see its own doc for the same pre-fetch + `extract-cache posting` pattern.

   **Checkpoint, repeated per batch.** Each posting's fetched content (often the largest single thing in context — full HTML job descriptions) is no longer needed once its row is written to "Fit Assessment." If working in batches of ~5 as Task 5 already recommends, **this is a checkpoint after every batch**, not just once at the end — the running total of "done this session" plus the remaining rows in `handoff_fit_assessments.json` is all that's needed to continue or to write the final summary.

6. **`python3 daily_pipeline.py build-review`** — the last script step, run once per day after all fit assessments are done (or at the end of each session if fit assessments are spanning multiple days). Syncs any Status/Skip Reason edits from `job-listings.xlsx` → "Shortlist" back to "All Listings" (the permanent record), then rebuilds "Shortlist" from scratch: all "All Listings" rows not yet marked Applied/Skipped, joined with Fit Assessment data, ranked best-first, capped at `review_list_size` (see `settings.md`). Prints the final run summary — retrieval counts, new companies vetted, unclassified titles/locations to review, and the top 3 Shortlist entries.

**Dependencies:** the script needs `pyyaml`, `python-dateutil`, and `openpyxl` (already used elsewhere in this pipeline). `requests` is no longer required for normal operation (kept only as an opt-in fallback for sandboxes where outbound HTTPS *does* work — see `fetch_greenhouse_jobs`/`fetch_greenhouse_job_detail` docstrings). If a package is missing: `pip install pyyaml python-dateutil --break-system-packages`.

**Why the checkpoints matter:** by design, every phase's important state lives in the spreadsheets and JSON handoff files, not in the conversation - that's what makes resuming an interrupted session safe (see "Resuming Between Sessions" below). The same property means compaction at these points loses nothing. If Cowork's context grows large and gets compacted automatically at some point mid-run anyway, calling out these boundaries explicitly makes it more likely that happens at one of them - a clean break with everything flushed to disk - rather than mid-fetch with partial state only in conversation history.

**End-of-run issues report:** both `phase1` and `phase2` end with either `No script issues.` or a bulleted list of every board-fetch/cache problem, capped runs, or unresolvable pending-revisit postings — detailed enough to hand back for a script fix. **Cowork should follow the same format for its own side of the run** (web_fetch failures, file permission errors, etc.) — at the end of a full run, report either "No issues." or a detailed list covering both the script's report and anything Cowork itself hit.

   **Checkpoint after `build-review`.** Everything is now in the spreadsheets. The run is done.

Only two things genuinely need LLM judgment:
- **Task 3's subagent** (Steps 1–4) — web research and synthesis, for the handful of companies (capped at `max_new_companies_per_run`) the triage selects.
- **Task 5's subagent** — reading actual job descriptions and weighing them against `candidate-profile.md`.

---

## Resuming Between Sessions

No separate checkpoint file. Each phase's own output **is** its checkpoint:

- **`phase1` re-run** → `seen-postings.xlsx` dedup means already-processed postings are skipped automatically; `company-tracker.xlsx` PENDING/Next-Review-Date logic means already-vetted companies are skipped too. Safe to re-run if interrupted.
- **`phase2` re-run** → needs `survivors.json` from the most recent `phase1` run. It dedups against existing `job-listings.xlsx` → "Shortlist" rows (by Job ID/URL), so re-running it doesn't create duplicates.
- **Task 5** → `job-listings.xlsx` → "Fit Assessment". `phase2` already excludes already-assessed rows when building `handoff_fit_assessments.json`.

**`phase1` and `phase2` are cheap** — no individual posting fetches beyond the board-list pulls. If a session gets cut off during either phase, or during the company-vetting step between them, just re-run from wherever it stopped — nothing gets double-processed.

**Task 5 is the expensive one** — one full posting fetch + analysis per Shortlist row — and is the only step that genuinely needs "have I done this one already?" checking (handled by `phase2`'s handoff). It's normal and expected for Task 5 to span multiple sessions, even multiple days, especially on a first run with a backlog.

**Report progress as each step finishes, not just at the end.** Both phases already print one line per task (counts in/out). If a session gets cut off partway through, these lines — visible in the chat transcript — are what tell you (and the next session) exactly where it stopped, without needing a separate log file.

---

## Reference: What `phase1` Implements (Tasks 0, 1, 2, Task 3 triage)

This section documents the logic `daily_pipeline.py phase1` implements — useful for understanding what happened in a run, or for modifying the script. It's not something Cowork needs to execute by hand.

### Task 0: Sync Applied Status

Check `job-listings.xlsx` ("Shortlist") for rows where `Applied?` has been marked (by B, since the last run). For each one, find the matching row in `seen-postings.xlsx` (by Job ID, falling back to Posting URL), and set `Applied? = Yes`, `Applied Date = today` (if not already set). This feeds the repost-detection check in Task 1.

---

### Task 1: Pull, Dedup & Title-Filter New Listings

**Pull from the companies listed in `company-boards.xlsx`** — for each row where Status = `Active` AND that row's Source is `Enabled` in `job-sources.xlsx`, the script reads that board's cache file (pre-fetched by Cowork — see "Running the Pipeline" above). Every currently-open job for that board is considered each run; **dedup against `seen-postings.xlsx` is what determines what's actually new**, not a date filter — Greenhouse's `?content=false` endpoint returns the full current job list every time, and `updated_at` reflects last *edit*, not first-seen date, so a date pre-filter would exclude most postings forever after day one.

**Volume cap:** stop processing once `settings.md` → "Max postings pulled per run" is reached, even if more boards remain. The script reports how many were left for next run — dedup means nothing is lost, just delayed.

**For Greenhouse boards specifically:** the cache file at `{cache_dir}/boards/{slug}.json` is `{"jobs": [...]}`, where each job has `id`, `title`, `absolute_url`, `location`, `updated_at`, etc. — written by `extract-cache board` from Cowork's `web_fetch` result, which also handles header-stripping and truncation (see "Running the Pipeline"). `absolute_url` is the Posting URL and should never be empty for a real Greenhouse job; if it somehow is, treat that single posting as broken (see below) rather than trying to reconstruct a URL after the fact.

**Job ID:** the JSON API's `id` field for each job (e.g. `7286376`) — store as text. This same number shows up in two other places you'll encounter: as `gh_jid=<id>` in a company's own branded careers-page URL (e.g. `stripe.com/jobs/search?gh_jid=7286376`), and as the trailing path segment on `job-boards.greenhouse.io`/`boards.greenhouse.io` URLs (`.../jobs/<id>`). All three are the same posting. This makes Job ID a more reliable cross-reference than the URL string itself — the same posting can legitimately have different-looking URLs depending on where it was pulled from, but the ID stays constant. If a future non-Greenhouse source is added, define its own Job ID extraction here (or leave blank if that source has no stable per-posting ID).

**Board URL formats:** `company-boards.xlsx` stores the human-facing `https://job-boards.greenhouse.io/{slug}` URL, while the API lives at `https://boards-api.greenhouse.io/v1/boards/{slug}/...`. `slug_from_board_url()` extracts `{slug}` correctly from either form (or from a `boards.greenhouse.io` URL) via a single regex — keep `company-boards.xlsx` in the human-facing format; no conversion needed.

**For every posting found, capture exactly these fields. If a field's value isn't shown on the board/listing page, record the literal string `Not Disclosed` — never leave it blank, and never skip a field:**

- Job Title
- Company Name — normalize per `settings.md` → "Company Name Normalization" (strip legal suffixes like "Inc.", "LLC") before recording. This keeps it an exact-match key everywhere downstream.
- Source (from `company-boards.xlsx`)
- Location
- Location Type — must be exactly one of `Remote`, `Hybrid`, `On-Site`, `Not Disclosed`. Map board phrasing accordingly: "Remote-USA"/"Remote, India"/etc. → `Remote`; "Hybrid" → `Hybrid`; a specific office location with no remote/hybrid language → `On-Site`; genuinely ambiguous → `Not Disclosed`.
- Posted Date — `Not Disclosed` if the board doesn't show one
- Pay Range — `Not Disclosed` if not shown (many postings legitimately omit this)
- Employment Type — must be exactly one of `Full-time`, `Contract`, `Part-time`, `Freelance`, `Temporary`, `Not Disclosed`. Map board phrasing accordingly: "Full Time"/"FT"/"Permanent" → `Full-time`; "Contractor"/"1099"/"C2C" → `Contract`; "Part Time"/"PT" → `Part-time`. If genuinely none of the five fit (e.g. an internship), use `Not Disclosed`.
- Posting URL
- Job ID — blank/`Not Disclosed` only if genuinely unobtainable (e.g. a non-Greenhouse source with no equivalent ID)

**Posting URL is the one exception to "Not Disclosed":** every posting has a canonical URL. If one can't be found for a specific posting, **don't add it to today's batch at all** — log it to `seen-postings.xlsx` → "Needs Link" (Company Name, Job Title, Source, Date Found) instead, and move on. It never reaches Task 2 onward. This is cheap to fix later (re-pull that company) and avoids a downstream task having to improvise a fix.

**Dedup against `seen-postings.xlsx` → "Seen Postings" (check before anything else):**

- **Job ID already present** (check this first — catches the same posting even if today's URL format differs from what was stored before) → update Last Seen Date, increment Times Seen. Already processed on a prior run — don't re-run the title filter or count it in today's batch. Skip.
- **No Job ID match, but Posting URL already present** → same handling as above. Fallback for postings where Job ID couldn't be determined.
- **Neither matches** → genuinely new. Before adding it, check: does an existing row with the same Company Name + Job Title have `Applied? = Yes`? If so, this looks like the same role reposted after you already applied — note `"possible repost after application"`. Either way, add a new row: Posting URL, Job ID, Company Name, Job Title, Source, First Seen Date = today, Last Seen Date = today, Times Seen = 1, Applied? = No, Notes = (repost flag if applicable). Then proceed to the title filter below.

**Location check (first content filter — runs before the title filter, for listings that passed dedup):**

- `Location Type = Remote` → passes automatically, regardless of `Location`.
- Otherwise, check `Location` against `location-filters.xlsx` → "Location Keywords" (case-insensitive substring match, any region). A match on **either** Sacramento Area or SF Bay Area keywords → passes.
- No match (and not Remote) → **bail**: don't run the title filter, don't proceed to Task 2. Update this posting's row in `seen-postings.xlsx` with `Notes = "Excluded - location (<Location value>)"` so it's recognized and skipped on future dedup passes without re-checking.

**Title filter (for listings that passed dedup *and* the location check):** check the Job Title against `job-title-filters.xlsx` → "Title Keywords" sheet, matching case-insensitive substrings, in this order:

1. **SKIP keywords** — if any match, discard entirely.
2. **CHECK keywords** — if any match (and no SKIP matched), Role Match = `CHECK`. Processed *after* PURSUE — lower priority.
3. **PURSUE keywords** — if any match (and nothing above matched), Role Match = `PURSUE`. Processed first.
4. **No match at all** — Role Match = `CHECK` (never silently dropped). Log to `job-title-filters.xlsx` → "Unclassified Titles" (Date Found, Job Title, Company Name), skipping if this exact title is already logged.

---

### Task 2: Group by Company

From the listings that survived Task 1 (Role Match = PURSUE or CHECK), extract the list of **unique** Company Names (already normalized per `settings.md` → "Company Name Normalization," so this is a plain exact-match dedup). For each, count how many surviving listings it has — this drives Task 3's priority order.

**Also add** any company in `company-tracker.xlsx` with `Status = PENDING` — these were deferred by the cap in a prior run and need another shot this run even if they have no new postings today. Treat these as highest priority (as if they had the most listings), so they're triaged first.

---

### Task 3 Triage: Resolving Companies Against `company-tracker.xlsx`

**Triage first — a single pass over `company-tracker.xlsx`, no subagent involved.** For each company from Task 2, match against `Company Tracker` → "Company Name" — both sides are already normalized per `settings.md`, so this is a plain exact match (case-insensitive). If a name still doesn't match after normalization, treat it as not found.

- **Not found** → needs vetting.
- **Found, `Status = BLACKLIST`** → resolved instantly. Task 4 will exclude its postings via Status. **Do not invoke the subagent.**
- **Found, `Status = PURSUE` or `WATCH`, and today is before `Next Review Date`** → resolved instantly, reuse this row as-is for Task 4. **Do not invoke the subagent.** Free — doesn't count against the cap. (The script computes `Next Review Date` itself from `Date Checked + Review Interval`, rather than relying on the cached `=EDATE(...)` formula value, which openpyxl can't read directly.)
- **Found, `Status = PENDING`, or `Next Review Date` has passed** → needs vetting.

**Cap:** of the companies needing vetting, take up to `settings.md` → "Max NEW companies vetted per run", prioritizing by the listing counts from Task 2 (most listings first — more riding on the verdict).

- **Within the cap** → goes into `handoff_companies.json` for `company-vetting-subagent.md` (Steps 1–4). The subagent no longer does its own tracker check — this triage already covered that — so it goes straight to research, and writes a new row to `company-tracker.xlsx` for companies that didn't have one.
- **Needs vetting but beyond the cap** → the script writes/updates the row directly, no subagent: `Status = PENDING`, `Company Fit`/`Ghost Verdict` blank, `Reason / Summary = "Awaiting vetting - capped this run"`, `Date Checked = today`, `Next Review Date = today` (so it's triaged first next run, ahead of routine re-checks — this is the queue that gradually builds out the avoid-list at a controlled pace).

**Why this matters:** most companies in a typical run will already be `BLACKLIST` or a recently-vetted `PURSUE`/`WATCH` — those resolve in the triage pass with zero subagent calls. The subagent (multiple web searches per company — the genuinely expensive part) only runs for the handful that need it. Previously, *every* company got a full subagent invocation just to determine which bucket it was in — almost certainly why a recent run stalled out after only getting through these per-company tracker checks.

---

## Reference: What `phase2` Implements (Task 4)

For each surviving listing from `survivors.json` (every one of which has a real Posting URL, per Task 1), look up its company's `Status` in `company-tracker.xlsx`:

- **Status = `BLACKLIST`** → discard.
- **Status = `WATCH`** → discard. Not a "revisit later" queue like `PENDING` — these listings just don't make it to today's Shortlist. If the company's status improves on a future `Next Review Date`, new postings from that point get evaluated fresh on their own.
- **Status = `PENDING`** → don't add to the shortlist yet — this company is queued for vetting next run (see Task 3). Add a note to this posting's row in `seen-postings.xlsx`: `"company pending - revisit"`. Once this company's `Status` resolves to `PURSUE`/`WATCH`/`BLACKLIST` in a future Task 3, that future run's Task 4 should also scan `seen-postings.xlsx` for rows with this note and this company, re-evaluate them (add to Shortlist if `PURSUE` and it clears the Company Fit floor below, just clear the note otherwise), and clear the note either way.
- **Status = `PURSUE`** → check `company-tracker.xlsx` → "Company Fit" against `settings.md` → "Shortlist Company Fit floor". `Strong Match` always passes. `Solid Contender` passes only if the floor is set to `Solid Contender` (the default). Anything passing is a candidate for the shortlist.

**Application Limit cap:** before adding candidates from a given company, check that company's `Application Limit` in `company-tracker.xlsx`.
- Blank/`[UNKNOWN]` → no cap, add all candidates.
- Set to N → add at most N from that company today, preferring Role Match = `PURSUE` over `CHECK`. The script reports any capped counts directly (e.g. "application limit: Stripe 6 matched, capped to 1").

The script adds surviving candidates to `job-listings.xlsx` ("All Listings"): Date Found, Job Title, Role Match, Company Name, Source, Location, Location Type, Pay Range, Employment Type, Posted Date, URL, Job ID, Company Fit, Notes, Status (blank — filled in via "Shortlist" sheet).

**Pending-revisit rescan:** also scans `seen-postings.xlsx` for postings noted `"company pending - revisit"` whose company has since resolved out of `PENDING` (possibly during this very run's vetting step). For `PURSUE` companies clearing the fit floor, it re-fetches the full posting via the Greenhouse API (Location/Pay/etc. aren't stored in `seen-postings.xlsx`) and adds it to the Shortlist; otherwise it just clears the note. (Role Match for these re-added postings defaults to `CHECK` — the original PURSUE/CHECK classification isn't persisted in `seen-postings.xlsx`.)

---

## Task 5: Run the Job Fit Assessment Subagent (per Shortlist entry)

**Before starting, read `candidate-profile.md` once** — don't re-embed its full contents into every sub-agent prompt if you spawn sub-agents; have them read the file themselves.

**For each Shortlist row, check `job-listings.xlsx` → "Fit Assessment" for an existing row matching that posting** — by Job ID first (more reliable across any URL differences), falling back to Posting URL if Job ID is blank. If found, skip — already assessed. Otherwise it's a candidate for this session's work.

**Prioritize:** among rows still needing assessment, work through `Company Fit = Strong Match` before `Solid Contender`, and within each, `Role Match = PURSUE` before `CHECK`.

**Keep each unit of work small.** Process at most `settings.md` → "Max Fit Assessments per run" postings this session (in batches of ~5 if using sub-agents, not 20+ — each posting fetch adds up in a sub-agent's context, and large batches are what caused a prior session to exhaust its usage mid-run). Write results to "Fit Assessment" as each batch returns — don't hold everything until the end. Anything beyond the cap carries over to the next session.

**Stop cleanly when usage feels tight** — finish the current batch, write its results, and stop. The next session re-checks "Fit Assessment" for missing Posting URLs and continues with whatever's left. A large first-run backlog (e.g. 100+ Shortlist rows from the initial company vetting) is expected to take several sessions; steady-state daily volume should be much smaller once the Shortlist itself is ~20/day.

Run the Job Fit Assessment Subagent (see `job-fit-assessment-subagent.md`) against `candidate-profile.md` for each posting processed this session, writing one row per posting to `job-listings.xlsx` ("Fit Assessment").

---

## End of Run

`phase1` and `phase2` each print their own one-line-per-task summary already (counts in/out, capped counts, etc.) — no need to reconstruct these by hand. After running both phases, just add a one-line Task 5 summary, e.g.:

> Task 5: 22 Shortlist rows needed assessment, 8 done this session (cap reached; 1 flagged with critical missing keywords, 1 flagged Pay Too Low), 14 remaining for next session.

If `handoff_companies.json` was empty (no vetting needed) or `handoff_fit_assessments.json` was empty (nothing new to assess), say so briefly — that's a normal, fast run.
