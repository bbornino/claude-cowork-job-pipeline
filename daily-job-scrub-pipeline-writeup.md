# Daily Job Scrub Pipeline (Python + Claude Cowork)

Point it at a list of company Greenhouse job boards and it pulls everything new, filters out the noise, vets any company you haven't seen before, and hands back a short, ranked list of postings actually worth applying to — each one scored against a real resume, not just keyword-matched.

It also took three days and never finished a clean run. That's the more interesting part of this writeup.

Source Code (GitHub): [https://github.com/bbornino/claude-cowork-job-pipeline](https://github.com/bbornino/claude-cowork-job-pipeline)

---

## A Quick Word on What's Happening Under the Hood

This isn't an AI app in the way the other projects in this series are. There's no single Anthropic API call doing the interesting work. Instead, there are two:

- **Company research** — given a company name, go figure out whether it's safe and worth pursuing: remote policy, layoffs/funding signals, LGBTQ+/DEI track record, AI-tooling culture, legal jurisdiction risk if HQ'd somewhere openly hostile to trans employees. This is retrieval-augmented synthesis — search, read, write a structured verdict.
- **Fit assessment** — given a job posting and a resume, score the match: which required skills are covered, which are gaps, does the pay range clear a floor, is the seniority right. This is structured comparison — read two documents, extract a judgment.

Everything else — fetching the boards, deduplicating against postings already seen, filtering by location and title, ranking what's left, capping how many postings per company make the final list, rotating which boards get priority on a given run — is plain deterministic Python. No model involved, no judgment calls, just data transformation. That split (~2,800 lines of Python, two narrow places where a model gets called) is itself the finding of this project, and it took building the wrong version first to see it clearly.

---

## What It Does

The pipeline runs in two phases:

**Phase 1** — pull every active board, skip anything already logged in `seen-postings.xlsx`, drop postings that don't match the configured location rules or that are older than a hard age cutoff (3 days by default — a real bug this project caught: month-old postings were quietly surviving a soft ranking penalty and getting recommended), classify titles as PURSUE/CHECK/SKIP, and write survivors to a tracking sheet.

**Phase 2** — for any company not yet vetted (or due for a 6-month re-check), run company research and write a verdict (PURSUE / WATCH / BLACKLIST) to `company-tracker.xlsx`. For surviving postings at PURSUE companies, run a fit assessment against the candidate's resume and build a ranked Shortlist, capped at 2 postings per company so one company with a lot of open roles can't crowd out everyone else.

Everything lives in plain `.xlsx` files — board list, company tracker, seen-postings log, the final Shortlist — so reviewing the output means opening a spreadsheet, not querying a database.

---

## Tools & Tech Stack

- **Core:** Python 3.12, openpyxl, PyYAML, requests
- **Execution environment (this iteration):** Claude Cowork — an agentic session driving bash, file edits, and web fetches
- **AI:** Anthropic Claude, invoked through two structured subagent prompts (`company-vetting-subagent.md`, `job-fit-assessment-subagent.md`) rather than direct API calls — see Build Notes for why that's the part being rebuilt
- **Data layer:** Excel workbooks (`.xlsx`) for every input and output — board config, company tracker, seen-postings log, Shortlist
- **Version control hygiene:** layered `.gitignore` plus a custom `snapshot.ps1` that hard-blocks personal-file patterns (`*.docx`, `*resume*`) and flags anything else unrecognized for manual confirmation before it can reach git history

---

## Build Notes — What Three Days Actually Bought

This is the section that matters more than the feature list.

The Python core came together fast and stayed solid the whole way through: fetch logic, deduplication, location/title filtering, ranking, a config file instead of hardcoded numbers, retry-with-backoff that correctly distinguishes transient failures (rate limits, timeouts) from permanent ones (a 403 that will never succeed no matter how many times you ask). None of that was the problem.

The problem was the execution environment. Four specific failures, each one real and each one costing real time:

**A FUSE-mounted sandbox serving stale reads.** Twice, the agent's own bash tool reported an older, shorter version of a file than what was verifiably on disk — confirmed by checking the file's exact byte count outside the session entirely. The agent's *other* tool (Read) could see the correct, current file at the same moment its bash tool couldn't. No fix exists short of bypassing the mount or restarting the session from scratch.

**A function that silently disappeared mid-refactor.** Splitting Greenhouse-specific logic into its own module, a function's section header and the code around it survived the move — the function body itself didn't. The gap sat invisible until the next live run threw an `AttributeError`, and the agent's first instinct was to blame the FUSE issue above rather than the actual missing code, which cost another round of debugging before the real cause surfaced.

**Self-repair that was worse than the failure.** When a cached data file looked corrupted, the agent's response was to reconstruct it from "known content" rather than stop and flag it — which meant quietly rewriting a real, populated configuration file as a blank one. Caught only because an independent copy happened to exist outside the session at the time.

**A real resume committed to git on the first attempt.** The safety check guarding against personal data leaking into version control matched an explicit list of filenames. A `.docx` resume sitting in the same folder wasn't on that list, so it sailed straight into the first commit. Fixed by adding pattern-based detection (`*.docx`, `*resume*`) and an allowlist that requires explicit confirmation for anything unrecognized — but the fact that the first version shipped without it is the actual lesson, not the fix.

None of these four were caused by wrong logic. They were caused by routing a deterministic, scriptable pipeline through a long-running conversational agent session that maintains its own state, its own caching, and its own judgment about how to recover when something looks wrong. Before assuming token cost was the limiting factor, the actual numbers were checked: a back-of-the-envelope estimate using this project's own measured output sizes put a 40-company, 50-fit-assessment run at under $2 in direct API cost. The constraint was never the budget. It was the shape of the tool versus the shape of the problem.

---

## Project Structure

```
claude-cowork-job-pipeline/
├── daily_pipeline.py              # Main script — phases, filtering, ranking, Shortlist build
├── greenhouse_api.py              # Greenhouse-specific fetch/parse, isolated for reuse with other ATSs later
├── discover_departments.py        # Standalone helper — maps a company's Greenhouse departments to IDs
├── company-vetting-subagent.md    # Structured research prompt — company safety/fit verdict
├── job-fit-assessment-subagent.md # Structured comparison prompt — resume vs. posting
├── settings.example.md            # All tunable caps and thresholds, with reasoning for each default
├── company-boards.example.xlsx    # Board config template
├── company-tracker.example.xlsx   # Vetting verdict template
├── snapshot.ps1                   # Git commit helper — layered personal-file detection
├── backup-data.ps1                # Local versioned backup for the real (gitignored) data files
└── README.md                      # Full setup + the honest retrospective this page is drawn from
```

---

## Setup

```bash
git clone https://github.com/bbornino/claude-cowork-job-pipeline
cd claude-cowork-job-pipeline

cp candidate-profile.example.md candidate-profile.md
cp settings.example.md settings.md
cp company-boards.example.xlsx company-boards.xlsx
cp company-tracker.example.xlsx company-tracker.xlsx
cp application-history.example.xlsx application-history.xlsx
# fill each one in with your own resume, targets, and thresholds

pip install pyyaml python-dateutil openpyxl --break-system-packages
python3 daily_pipeline.py phase1
```

The README covers the rest — department filtering setup, the two AI-dependent steps, and the full retrospective above in more detail.

---

## What's Not Built Yet

- **The actual rebuild.** Same Python core, same two subagent prompts — reimplemented as direct Anthropic API calls from a standalone script instead of Cowork-orchestrated steps. No session, no sandbox, no agent maintaining state across dozens of tool calls. This is the next project in practice, not just a stated intention.
- **Scheduling.** Right now this runs on demand. A cron job or scheduled task is the obvious next step once the standalone version exists.
- **Multi-ATS support.** `greenhouse_api.py` was deliberately split out from the main script so a Lever or Workday module could sit next to it later — not done yet.

---

## Background

Built across three days as a personal tool during an active job search, with the explicit goal of cutting down the ATS-screening problem that comes with high-volume cold applications. It didn't end up automating the job search. It ended up teaching a sharper lesson about where agentic tooling actually earns its keep: company research and resume-matching are exactly the kind of unstructured-reading, structured-judgment work a language model is good at. Fetching, filtering, deduplicating, and ranking spreadsheet rows is exactly the kind of work it shouldn't be doing turn-by-turn in an interactive session. The next version applies that lesson directly instead of relearning it.
