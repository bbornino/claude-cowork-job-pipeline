# Daily Job Scrub Pipeline

> **Status: paused, not abandoned.** The Python pipeline itself (`daily_pipeline.py` / `greenhouse_api.py`) is tested and works. Running it *through Cowork's interactive session*, the way this repo was originally set up to do, never completed a clean end-to-end run at even small scale (8 companies, 15 fit assessments) across three days of active building. That's the actual finding here, and it's documented in detail below rather than smoothed over — see "What This Project Actually Demonstrated" near the bottom. The deterministic 95% of this pipeline (fetch, dedupe, filter, rank) is correct and reusable; the next iteration moves the two genuinely AI-dependent steps (company research, resume-to-posting fit) into direct Anthropic API calls from a standalone script, with no agentic session in the loop at all.

A daily-run pipeline that pulls new job postings from a curated list of company job boards (currently Greenhouse), filters out anything irrelevant, vets new companies, and builds a short, prioritized list of postings worth actually applying to — plus a per-posting fit assessment against your resume.

Most of it is a script. The two genuinely "smart" steps (researching a new company, and reading a job description closely) are done by Claude/Cowork.

---

## Before You Start: Set Cowork's Mode to "Act Without Asking"

By default, Cowork pauses for approval before each action (including `web_fetch`) — this is what causes the "allow getting this page from Greenhouse?" prompts. For a walk-away run, switch the mode selector on Cowork's chat input from **"Ask before acting"** to **"Act without asking"** before sending the task below.

Tradeoff to be aware of: this applies to *all* web access during the session — not just the repeated Greenhouse API fetches, but also whatever broader research `company-vetting-subagent.md` does for a new company (Glassdoor, LinkedIn, company sites, etc.). Anthropic calls this a higher-risk mode for that reason. For this pipeline — mostly one well-known read-only JSON API plus mainstream company-research sites — that seems like a reasonable tradeoff for an unattended run, but it's your call. Cowork still asks before *deleting* files in either mode, which doesn't come up here since the pipeline never deletes anything.

---

## One-Time Setup: Department Filtering

By default, the pipeline pulls every job from every board — e.g. Anthropic alone returns 139 listings, most of them Sales, Recruiting, Legal, etc. that get filtered out downstream anyway, but only after burning tokens on title/location checks. Department filtering fetches just the relevant departments (Engineering, IT) at the source, which is both cheaper and cleaner.

This is a one-time setup per company (department lists rarely change), not part of the daily run.

**Scope: only companies with no verdict yet, or `PENDING`/`PURSUE`.** Never run this for a company already `BLACKLIST` or `WATCH` in `company-tracker.xlsx` — there's no reason to spend a fetch finding departments for a company you've already decided against. Cross-reference `company-boards.xlsx` against `company-tracker.xlsx` yourself before starting; don't just loop over every row in `company-boards.xlsx`.

---BEGIN DEPARTMENT DISCOVERY TASK---

First, check `company-tracker.xlsx` and list which companies in `company-boards.xlsx` are NOT already `BLACKLIST` or `WATCH` (i.e. have no tracker row yet, or are `PENDING`/`PURSUE`). Show me that list before fetching anything, so I can confirm the scope is right.

Then run the batch command directly on the bare slugs — do NOT web_fetch anything yet:

```
python3 daily_pipeline.py list-departments-batch slug1 slug2 slug3 ...
```

This tries a direct Python HTTP call for each company first, with no response-size limit at all (unlike `web_fetch`, which caps around 94KB and is why this endpoint kept getting truncated before). If direct access works from this environment, you get the complete department list with zero truncation risk and nothing else to do.

The report will end with a list of any companies where the direct call failed and no cache file existed either — only for THOSE specific companies, `web_fetch` `https://boards-api.greenhouse.io/v1/boards/{slug}/departments`, save the result to a file, then re-run the batch command for just those, using `slug=file` instead of the bare slug:

```
python3 daily_pipeline.py list-departments-batch slug1=file1.txt slug2=file2.txt
```

This second command also writes the combined report to `department_discovery_report.txt` in case the chat output gets cut off. Show me the complete output — don't filter or summarize it, I want to see the full list myself before deciding which departments to keep.

If a `web_fetch`-sourced file was itself too large and got truncated, the report will say so per-company and mark job counts with a `?` — those counts are unreliable, but department IDs and names are still extracted correctly in most cases, since they appear before each department's job list in the response. If a company's list looks suspiciously short or cuts off mid-alphabet, flag it — that company may need a second look (e.g. a different/larger fetch) rather than treating the partial list as complete.

**If the same companies keep coming back truncated no matter what** (i.e. `/departments` itself is too large to ever fetch cleanly for that board, with or without direct access), switch to the lighter alternative for just those companies:

```
python3 daily_pipeline.py derive-departments-batch slug1 slug2 ...
```

This skips `/departments` entirely and instead fetches the much smaller `/jobs?content=false` (the same endpoint the daily pipeline already uses) and derives department id/name/job-count directly from the jobs returned, rather than from Greenhouse's separate department registry. Same bare-slug-tries-direct-first behavior and `slug=file` fallback as `list-departments-batch`. Two real differences worth knowing: job counts here are exact, not `?`-marked, since they're counted from jobs actually in hand rather than recovered from a cut-off response; and a department with zero current postings won't show up at all (there's no job to derive it from) — which is fine for this purpose, since an empty department isn't useful for the `Departments` filter anyway.

---END DEPARTMENT DISCOVERY TASK---

Once you have the full list, pick the IDs for Engineering/IT-equivalent departments (names vary by company — could be "Engineering," "Product & Engineering," "R&D," "Infrastructure," "Data," "IT," etc.) and add them as a comma-separated list to the `Departments` column in `company-boards.xlsx` for that company's row. Leave `Departments` blank for any company to keep pulling everything (no filter).

Once `Departments` is filled in for a company, the daily task automatically fetches only those departments for it — see step 1 below, which checks for this.

**Resuming a partial discovery session.** If a previous attempt fetched some or all of the raw department responses but didn't finish the batch report (e.g. it ran out of context mid-way), don't refetch from scratch — reuse what's already saved. Use this instead:

---BEGIN DEPARTMENT DISCOVERY RESUME TASK---

If this is a fresh session (no `/tmp/pipeline_data/` yet, or `$COWORK_DIR` isn't set), run `bash setup.sh` first as usual.

Run the batch command on the bare slugs for every in-scope company again (cross-referenced against `company-tracker.xlsx` as above — skip BLACKLIST/WATCH) — this re-attempts the direct Python fetch, which is cheap and fast, so there's no real cost to retrying it even for companies that succeeded before:

```
python3 daily_pipeline.py list-departments-batch slug1 slug2 ...
```

For any companies the report says still need a manual fetch: check first whether a raw file from a prior session already exists (`ls /tmp/pipeline_data/*depts* /tmp/pipeline_data/*departments* 2>/dev/null`, and check the current working directory too) before doing a fresh `web_fetch` — reuse what's already saved rather than refetching. Then re-run the batch command with `slug=file` for just those.

Once the report is complete, immediately copy `department_discovery_report.txt` back to the Cowork folder — don't wait until the end of the session:

```
cp /tmp/pipeline_data/department_discovery_report.txt "$COWORK_DIR/"
```

Show me the complete report output in chat as well.

---END DEPARTMENT DISCOVERY RESUME TASK---

---

## Running It (Cowork Task)

Copy everything between the lines below and paste it into Cowork:

---BEGIN COWORK TASK---

Run today's daily job scrub pipeline, following `daily-job-scrub-pipeline.md`. Work entirely out of `/tmp/pipeline_data/` to avoid FUSE mount stale-cache issues — the bash sandbox may not see files edited via the Read/Write tools unless they're in `/tmp`.

Setup (run once at the start of each session):

```
bash setup.sh
```

This finds the Cowork folder, copies everything to `/tmp/pipeline_data/`, and verifies `daily_pipeline.py` compiles, has the right line count, and successfully imports alongside `greenhouse_api.py` (the module that holds all Greenhouse-specific fetch/parse logic — both files need to be present and consistent). If it reports a warning, stop and report before proceeding. Run all subsequent commands from `/tmp/pipeline_data/`.

Then run the pipeline from `/tmp/pipeline_data/`:

1. Run `python3 daily_pipeline.py phase1` directly — don't `web_fetch` anything yet. Phase1 tries a direct Python HTTP call for every active board first (no response-size cap, no truncation risk), falling back to a cache file only for boards where that fails. Its "Task 1" summary line reports the fetch source breakdown (e.g. "fetch source: 12 direct, 4 none") and lists any board errors with the exact `web_fetch` URL and `extract-cache` command needed. For each board listed as a failure: `web_fetch` that URL (the script's error message gives the exact one — the full job list if `Departments` is blank for that board, or each configured department's URL if not), save the result, then process them all in one call instead of one at a time:

   ```
   python3 daily_pipeline.py extract-cache-batch board:cloudflare:cloudflare.txt dept:gitlab:4115236002:gitlab_ai.txt dept:gitlab:4135580002:gitlab_arch.txt ...
   ```

   (Each spec is `board:slug:file`, `dept:slug:dept_id:file`, or `posting:job_id:file` — colon-separated, one per failed fetch. A department-filtered board needs one spec per configured department ID, so this matters most for boards with several departments — Reddit has 11, MongoDB has 13.) Then re-run `python3 daily_pipeline.py phase1` again. This second run skips boards that already succeeded and only needs the ones you just cached.

   **PURSUE companies get processed before WATCH/unvetted ones, every run, automatically** — Task 1 prioritizes boards whose company is `PURSUE` in `company-tracker.xlsx`, and only spends remaining `max_postings_per_run` budget on the rest. Within the PURSUE group, a persisted rotation (`board_rotation_state.json`, gitignored, regenerates if deleted) makes sure each PURSUE board gets first crack at the cap on a rotating basis across runs, rather than the same handful of boards (whichever come first in `company-boards.xlsx`) winning every single time. The Task 1 summary's "PURSUE-board rotation" line shows which board started this run and which one starts next.
2. Run `python3 daily_pipeline.py phase1` and report its summary.

Checkpoint: everything from steps 1-2 is now saved to disk. Compact context here if you can — only the printed summary and `handoff_companies.json` matter going forward.

3. For each company in `handoff_companies.json` (if any), run `company-vetting-subagent.md` and write the verdict to `company-tracker.xlsx` in `/tmp/pipeline_data/`.

Checkpoint: verdicts are now in `company-tracker.xlsx`. Compact here too if you can — the research behind each verdict doesn't need to be retained.

4. Run `python3 daily_pipeline.py phase2` and report its summary.

Checkpoint: results are now saved to disk. Compact again if you can — only `handoff_fit_assessments.json` matters going forward.

5. For each posting in `handoff_fit_assessments.json` (if any): pre-fetch via `web_fetch` `https://boards-api.greenhouse.io/v1/boards/{slug}/jobs/{Job ID}?content=true` + `python3 daily_pipeline.py extract-cache posting {Job ID} <result_file>`, then run `job-fit-assessment-subagent.md` and write the result to `job-listings.xlsx` ("Fit Assessment") in `/tmp/pipeline_data/`. Work in batches of ~5; after each batch is written, compact context if you can.
6. Run `python3 daily_pipeline.py build-review` and show its complete output.
7. Copy all modified files back to the Cowork folder using the `cp` command printed at the end of `setup.sh` output.
8. End with an issues report: either "No issues." or a detailed bulleted list covering (a) anything `phase1`/`phase2` printed under "Script issues" — copy verbatim, don't summarize — and (b) any Cowork-side problems (web_fetch failures, file permission errors, FUSE issues, anything that needed a workaround), in enough detail to fix the scripts or docs.

Report progress after each step. If usage runs low, finish the current step or batch, write its results, copy back to the Cowork folder, and stop — say where you stopped so the next session can pick up.

---END COWORK TASK---



**After `phase1` (and again after `phase2`), check the folder for these files** — useful if anything looked off (errors, unexpected counts) and you want a second opinion:

- `phase1_report.json` — full Task 0-3 counts, including the *complete, untruncated* text of any board-fetch errors (HTTP status, deny-reason header, or a short response-body snippet)
- `phase2_report.json` — full Task 4 counts (after `phase2`)
- `handoff_companies.json` / `handoff_fit_assessments.json` — what got queued for the subagent steps

These are small (typically under 1KB, even with several errors) and live directly in the local folder Cowork is working in — **just upload them to your Claude.ai conversation** to get a second opinion on a run. No need to ask Cowork to print/paste their contents.

That's it — the rest of this file is reference material for understanding or tuning the pipeline, not something you need to read to run it day-to-day.

---

## Resetting for a Fresh Start

Run `reset.bat` (double-click, or `reset` from a command prompt in the folder) to wipe stale data and start clean. Takes about 2 seconds.

**What it resets** (clears data, keeps headers): `seen-postings.xlsx`, `job-listings.xlsx`

**What it deletes** (ephemeral files): `run-log.xlsx`, `survivors.json`, `handoff_*.json`, `phase*_report.json`

**What it never touches**: `company-tracker.xlsx`, `application-history.xlsx`, `company-boards.xlsx`, `job-sources.xlsx`, `job-title-filters.xlsx`, `location-filters.xlsx`, `candidate-profile.md`, `settings.md`, the pipeline docs, and `daily_pipeline.py` — all your real configuration stays intact.

The template files (`seen-postings-template.xlsx`, `job-listings-template.xlsx`) are what `reset.bat` copies from. Don't edit or delete them.

---

## After the Run

**One file to open: `job-listings.xlsx` → "Shortlist" tab** (first tab, opens by default).

This is rebuilt fresh every run — ranked best-first, capped at `review_list_size` (default 40, set in `settings.md`). The top entries are your best candidates based on Company Fit, then Role Match, then (for assessed rows) Requirements Met ratio, Pay Flag, and ATS critical gaps. Only listings you haven't already acted on appear here.

For each one: set **Status** = `Applied` or `Skipped`. Add a brief **Skip Reason** if you want a record of why (e.g. "too much Java", "comp too low"). That's it — changes sync back to the permanent "All Listings" tab automatically on the next run, and Skipped/Applied rows drop off the Shortlist.

**Three other things `build-review` may flag:**

- **New companies vetted** → check `company-tracker.xlsx`. Review each new row's verdict and adjust Status/Company Fit if you disagree with Cowork's research.
- **Unclassified job titles** → check `job-title-filters.xlsx` → "Unclassified Titles". Titles that didn't match any keyword rule landed here as CHECK by default. Add each to "Title Keywords" as SKIP/PURSUE/CHECK so they're classified correctly next run.
- **Unrecognized locations** → check `location-filters.xlsx` → "Unrecognized Locations". Location strings that passed via a keyword substring (e.g. "South San Francisco, CA" matched via "san francisco") but haven't been explicitly listed. Add ones you want to keep to "Location Keywords"; future listings with that exact string won't be re-flagged.

None of these are urgent — they accumulate as the pipeline runs and you can batch-review them whenever. The Shortlist is what matters daily.

---

## Saving Your Work: `snapshot.ps1`

If this repo is under git (see the tarball/`.gitignore` setup for the public-safe version), run `.\snapshot.ps1` after anything you'd be annoyed to redo — a clean daily run, finishing department discovery for a new company, before trying anything experimental in a Cowork session. It's a deliberate, manual checkpoint: shows you what changed, refuses to commit any of the personal files in `.gitignore` even if something staged them by mistake, commits locally with a timestamp, and stops there — no auto-push, nothing scheduled, nothing running without you typing the command.

This isn't wired into the daily pipeline on purpose. After running into a Cowork session that silently reconstructed a binary file from guessed content when the real one looked corrupted, the goal is a save mechanism that's simple enough to fully trust — adding automatic commits inside the same automation that just had a near-miss would be solving one reliability problem by adding another moving part. Run it yourself, often, especially right after anything that worked.

```
.\snapshot.ps1                      # auto-generated commit message
.\snapshot.ps1 "filled in MongoDB departments"   # custom message
```

**For the data itself** (everything `.gitignore` excludes — your resume, comp targets, target companies, vetting verdicts, application history, live scrape results), use `.\backup-data.ps1` instead. This isn't git-based at all — it just copies each file into a local `backup\` folder, with simple versioned rotation: the most recent backup is always the unnumbered `<name>.backup.xlsx`, and each backup run before that bumps up a number (`.backup.1.xlsx`, `.backup.2.xlsx`, ...), keeping the last 5 versions per file by default. `backup\` itself is gitignored (the script adds the line for you if it's somehow missing) — these are local safety copies, not something that belongs in version control alongside the personal originals they're duplicating.

```
.\backup-data.ps1
```

Run both after a clean run or before anything risky — `snapshot.ps1` for the code, `backup-data.ps1` for the data. They're separate because the things they protect are fundamentally different: one is meant to live in shared history, the other is explicitly meant to never leave your machine.

**First run of either script will likely fail with `UnauthorizedAccess` / "is not digitally signed."** This is Windows blocking unsigned `.ps1` files by default, and it gets worse if the file was downloaded (Windows tags downloaded files with a "Mark of the Web" flag, which `RemoteSigned` still blocks even after changing the execution policy). Fix both layers once, then it's solved permanently for any `.ps1` file in this repo:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
Get-Item *.ps1 | Unblock-File
```

If it still fails after that, check whether the Mark of the Web is actually present (`Get-Item .\snapshot.ps1 -Stream Zone.Identifier` — an error means it's already clear) and re-run `Unblock-File` on the specific script. As a one-off workaround that doesn't need either step: `powershell -ExecutionPolicy Bypass -File .\snapshot.ps1`.

---

## What a Run Actually Does

1. **Pre-fetch** (Cowork) — fetches each active company board's job list (and, rarely, individual posting details) via `web_fetch`, normalized into a local cache by `extract-cache`.
2. **`phase1`** (script) — reads the cache, dedups against postings already seen, drops anything in the wrong location or with a SKIP-listed title, and figures out which companies (if any) need fresh research.
3. **Company vetting** (Claude, only for companies `phase1` flags — usually 0–8 per day) — researches a new or stale company and records a verdict (PURSUE / WATCH / BLACKLIST + a fit rating) in `company-tracker.xlsx`.
4. **`phase2`** (script) — using the now-current company verdicts, decides which surviving postings make today's Shortlist, and picks which Shortlist rows (if any) need a closer look.
5. **Fit assessment** (Claude, only for postings `phase2` flags — usually 5–10 per day) — pre-fetches each posting's full content the same way, then reads it and compares against your resume, noting gaps and ATS keywords.

Steps 3 and 5's analysis are the only ones that need Claude's judgment. Everything else is fetching (Cowork's `web_fetch`, since the script can't make outbound calls itself) plus the script reading and writing the spreadsheets below.

---

## File Guide

**Run this:**
- `daily_pipeline.py` — the script. `phase1` / `phase2` as described above.

**How-to / reference docs:**
- `daily-job-scrub-pipeline.md` — the full spec: what each phase does, how to resume an interrupted session, what the "End of Run" report should look like.
- `company-vetting-subagent.md` — instructions Claude follows to research one company (used in step 2).
- `job-fit-assessment-subagent.md` — instructions Claude follows to assess one posting (used in step 4).
- `candidate-profile.md` — your resume/skills summary, used as the baseline for both vetting (role fit) and fit assessment (requirements/ATS keywords).
- `settings.md` — every tunable number (daily caps, pay thresholds, color coding) plus a config block the script reads directly. **This is the first place to look if you want to change how the pipeline behaves.**

**Data files (the script reads and writes these — you mostly just review them):**
- `company-boards.xlsx` — which companies' job boards to pull from, and where. `Departments` column (optional) restricts a board to specific Greenhouse department IDs — see "One-Time Setup: Department Filtering" above.
- `job-sources.xlsx` — which ATS platforms (Greenhouse, etc.) are turned on.
- `job-title-filters.xlsx` — keyword rules for what titles to skip/flag/pursue, plus a log of titles that didn't match any rule (worth reviewing occasionally).
- `location-filters.xlsx` — which locations count as "close enough."
- `seen-postings.xlsx` — every posting ever pulled, for dedup. Mark `Applied?` = Yes here (or on the Shortlist) once you apply.
- `company-tracker.xlsx` — one row per company researched so far, with its verdict and a re-check date.
- `job-listings.xlsx` — **the actual output.** Three tabs:
  - **"Shortlist"** (opens first) — the ~40 best listings you haven't acted on yet, ranked and rebuilt fresh every run by `build-review`. This is what you look at daily. Set Status = Applied or Skipped here.
  - **"All Listings"** — the permanent record of every posting that passed the filters. `build-review` reads this; `phase2` writes to it. You rarely need to open this tab directly.
  - **"Fit Assessment"** — per-posting fit detail (requirements met, ATS keywords, pay flag) for assessed listings. Written by `job-fit-assessment-subagent.md`; joined into "Shortlist" automatically.

**Generated each run (safe to ignore/delete — regenerated by `phase1`/`phase2`):**
- `survivors.json`, `handoff_companies.json`, `handoff_fit_assessments.json`
- `phase1_report.json`, `phase2_report.json` — full diagnostics for the commands above
- `/tmp/pipeline_cache/boards/*.json`, `/tmp/pipeline_cache/postings/*.json` — pre-fetched board/posting data, written by `extract-cache`. Lives outside the project folder (Linux ephemeral filesystem) on purpose — see "First-Run / Tuning Notes."

**Deprecated:**
- `run-log.xlsx` — from an earlier design; no longer used. The output files themselves now serve as the resumability checkpoint (see "Resuming Between Sessions" in the pipeline doc). Safe to delete.

---

## First-Run / Tuning Notes

- All the daily caps (postings pulled, companies vetted, fit assessments) live in the YAML block at the top of `settings.md`. They're set conservatively for testing — raise them once a run completes comfortably within Cowork's usage limits.
- The very first run will likely have a large Shortlist backlog (company vetting starts from scratch). Expect steps 3 and 5 to span several sessions at first; steady-state daily volume should be much smaller.
- `daily_pipeline.py` needs `pyyaml`, `python-dateutil`, and `openpyxl`. If any are missing: `pip install pyyaml python-dateutil openpyxl --break-system-packages`.
- **Confirmed:** outbound HTTPS from Cowork's bash sandbox is blocked (403 at the egress proxy, not Greenhouse-side). All fetching is done by Cowork's `web_fetch` + `extract-cache`, per the task above — the script never calls the network itself by default.
- **Cache directory:** defaults to `/tmp/pipeline_cache` (Linux ephemeral filesystem), deliberately *outside* the project folder. If the project folder is on a Windows mount, files created/copied there this session can't be overwritten or deleted from the bash sandbox (`PermissionError [Errno 13]` / `rm: Operation not permitted`) — only new files can be created. `/tmp` avoids this entirely. Don't redirect `--cache-dir`/`PIPELINE_CACHE_DIR` onto the project folder unless you've confirmed overwrites work there.
- **If `daily_pipeline.py` looks truncated or won't compile** after being copied/transferred into Cowork's folder, re-copy it from this conversation's output and verify with `python3 -m py_compile daily_pipeline.py` — it should end with the `if __name__ == "__main__":` block dispatching `phase1`/`phase2`/`extract-cache`.

---

## What This Project Actually Demonstrated

Three days, easily 4-5 active hours each, building and rebuilding this through Claude Cowork. Zero clean end-to-end runs — at a scale (8 companies, 15 fit assessments) small enough that token cost was never the actual constraint. A back-of-the-envelope estimate for the same work done as direct API calls instead, using real measured output sizes from this project's own data, comes out under $2 for 40 companies and 50 fit assessments. So the failures here were never about budget. They were about running deterministic, scriptable work through an interactive agent session built for open-ended reasoning.

The specific failure modes, in case they're useful to someone hitting the same wall:

- **A FUSE-mounted sandbox serving stale file reads.** Twice, the bash tool inside the session reported a smaller, older version of a file than what was verifiably on disk (confirmed via the file's own Properties dialog on Windows — exact byte-count mismatch). No root-cause fix found; only workaround was bypassing the mount via a different tool (Read instead of bash) or restarting the session entirely.
- **A missing function signature, silently surviving a refactor.** During a module extraction (pulling Greenhouse-specific logic into its own file), a function's section header and its surrounding code both got carried over correctly — the function body itself didn't. The bug sat invisible until the very next live run crashed on it, and Cowork's own diagnosis of the crash initially blamed FUSE staleness rather than the actual missing code.
- **Self-repair worse than the failure it was repairing.** When a cached data file looked corrupted, the session's instinct was to reconstruct it from "known content" rather than stop and report the corruption — which meant guessing at a blank version of a file that, in reality, held real configuration data. Caught only because an independent copy existed outside the session.
- **A personal file (a real resume, `.docx`) committed to git history** on the first attempt, because the safety check only matched an explicit list of filenames rather than file *type* or pattern. Fixed with a layered check (exact names, then patterns like `*.docx`/`*resume*`, then an allowlist requiring explicit confirmation for anything unrecognized) — but the first version shipping without that layering is itself the lesson.

None of these were caused by the underlying logic being wrong. They were caused by routing plain control flow — fetch, dedupe, filter, rank, cap, write — through a conversational tool-calling loop that has its own session state, its own caching layer, and its own judgment calls about how to recover from confusion. That's a mismatch between the shape of the problem and the shape of the tool, not a flaw in the pipeline design.

**What carries forward:** `daily_pipeline.py` and `greenhouse_api.py` are correct, tested, and architecture-agnostic — the same fetch/filter/rank logic works identically whether it's invoked from a Cowork session or a plain script. The two genuinely AI-dependent steps (company research, resume-to-posting fit assessment) are exactly the kind of work — reading unstructured text, producing structured judgment — that a direct Anthropic API call handles well without any of the session-state failure modes above. That's the next iteration: same Python core, the two smart steps moved to direct `anthropic` API calls (likely Sonnet, given the cost math), no Cowork session in the loop at all.
