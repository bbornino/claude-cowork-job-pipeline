"""
discover_departments.py — run this from your own machine, not from Cowork.

WHY THIS EXISTS
-----------------
Cowork's sandbox has outbound HTTPS blocked at the network/proxy level
(confirmed repeatedly: every direct Python fetch attempt gets a 403,
regardless of which Greenhouse endpoint). Cowork's own web_fetch tool works
around that, but caps responses around 94KB, which keeps cutting off
Greenhouse's department/job-list responses for the larger boards
(Anthropic, Reddit, Databricks, MongoDB, GitLab) before they finish.

Your own machine has no such restriction. This script makes the exact same
HTTP calls directly, with no fetch-size cap at all, then writes the same
department_discovery_report.txt format the rest of the pipeline already
expects — so the output slots in exactly where Cowork's batch commands
would have put it.

WHAT IT DOES
-------------
For each company slug given:
  1. Fetches https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=false
     (the full job list, no per-job description - light payload)
  2. Derives department id/name/job-count directly from the jobs returned,
     rather than fetching the heavier /departments endpoint (which embeds
     full job content per posting and is the thing that kept truncating).

USAGE
------
    python discover_departments.py gitlab asana figma anthropic amplitude databricks reddit datadog mongodb

Or edit the SLUGS list below and run with no arguments:
    python discover_departments.py

OUTPUT
-------
Writes department_discovery_report.txt in the current directory, in the
same format the pipeline's derive-departments-batch command produces.
Copy that file into your Cowork folder (or just read it yourself) and use
it the same way: pick the Engineering/IT department IDs per company and
add them to the Departments column in company-boards.xlsx.

REQUIREMENTS
-------------
Python 3.7+ and the `requests` library:
    pip install requests
"""

import json
import sys
import time
import warnings

# Suppresses RequestsDependencyWarning ("urllib3 ... doesn't match a
# supported version") - cosmetic only, doesn't affect correctness, but
# noisy on every run if your urllib3/chardet versions are slightly out of
# requests' expected range. The real fix is `pip install --upgrade requests
# urllib3 charset_normalizer`, but this keeps the warning out of the way
# either way.
warnings.filterwarnings("ignore", message=".*doesn't match a supported version.*")

try:
    import requests
except ImportError:
    print("ERROR: the 'requests' library is required.")
    print("Install it with: pip install requests")
    sys.exit(1)


# Edit this list if you'd rather not pass slugs on the command line.
SLUGS = [
    "gitlab", "asana", "figma", "anthropic", "amplitude",
    "databricks", "reddit", "datadog", "mongodb",
]

TIMEOUT_SECONDS = 30  # generous - no fetch-size cap here, so a slow response
                       # finishing late is fine, unlike Cowork's web_fetch.


def jobs_url(slug):
    # content=true, not content=false: Greenhouse only includes the
    # `departments` array (along with `offices` and the full description)
    # when content=true. With content=false, departments is omitted
    # entirely - which is why the first version of this script got 0
    # departments back for every company despite fetching successfully.
    # This script has no fetch-size cap (unlike Cowork's web_fetch), so the
    # extra per-job description weight from content=true costs nothing here.
    return f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=true"


def fetch_jobs(slug):
    """Returns (jobs_list, error_or_None). Never raises - every failure
    mode is captured and returned as a short error string instead, same
    convention as the rest of the pipeline."""
    url = jobs_url(slug)
    try:
        resp = requests.get(url, timeout=TIMEOUT_SECONDS, headers={"User-Agent": "Mozilla/5.0"})
    except Exception as e:
        return [], f"{type(e).__name__}: {e}"
    if resp.status_code != 200:
        snippet = (resp.text or "")[:200]
        return [], f"HTTP {resp.status_code} from {url} - {snippet}"
    try:
        return resp.json().get("jobs", []), None
    except ValueError as e:
        return [], f"Invalid JSON: {e} - body started with: {resp.text[:200]!r}"


def derive_departments_from_jobs(jobs):
    """Same logic as greenhouse_api.py's derive_departments_from_jobs -
    duplicated here so this script has zero dependency on the pipeline
    codebase and can be run standalone from any machine with Python."""
    counts = {}
    for job in jobs:
        for d in job.get("departments") or []:
            did = d.get("id")
            name = d.get("name")
            if did is None:
                continue
            if did not in counts:
                counts[did] = {"id": did, "name": name, "count": 0}
            counts[did]["count"] += 1
    return sorted(counts.values(), key=lambda x: (x["name"] or ""))


def main():
    slugs = sys.argv[1:] if len(sys.argv) > 1 else SLUGS
    if not slugs:
        print("No slugs given. Pass them as arguments or edit the SLUGS list in this file.")
        sys.exit(1)

    print(f"Fetching department data for {len(slugs)} company board(s)...")
    print("(This calls Greenhouse directly - no Cowork, no fetch-size cap.)\n")

    lines = ["DEPARTMENT DISCOVERY REPORT (derived from /jobs, fetched directly - no Cowork)",
             "=" * 60, ""]
    any_warning = False

    for i, slug in enumerate(slugs, start=1):
        print(f"  [{i}/{len(slugs)}] {slug}...", end=" ", flush=True)
        jobs, err = fetch_jobs(slug)
        if err:
            print(f"FAILED: {err}")
            lines.append(f"\n--- {slug} ---")
            lines.append(f"ERROR: {err}")
            any_warning = True
            time.sleep(0.5)
            continue

        depts = derive_departments_from_jobs(jobs)
        print(f"OK ({len(jobs)} jobs, {len(depts)} departments)")
        lines.append(f"\n--- {slug} --- ({len(jobs)} job(s) retrieved, complete - no truncation)")
        if not depts:
            lines.append("No departments found in the retrieved jobs.")
            continue
        lines.append(f"{len(depts)} department(s) seen in these jobs:")
        lines.append(f"{'ID':>10}  {'Jobs':>5}  Name")
        lines.append("-" * 60)
        for d in depts:
            lines.append(f"{d['id']:>10}  {d['count']:>5}  {d['name']}")

        time.sleep(0.5)  # light courtesy delay between requests - not required,
                          # just polite when hitting the same host repeatedly.

    lines.append("")
    lines.append("=" * 60)
    if any_warning:
        lines.append("Some boards failed to fetch entirely - see ERROR lines above.")
        lines.append("These are real connectivity/API errors, not truncation - this script")
        lines.append("has no response-size cap, so anything that succeeded is complete.")
    lines.append("")
    lines.append("Add desired IDs (comma-separated) to the 'Departments' column in")
    lines.append("company-boards.xlsx for each company. Leave blank to keep pulling")
    lines.append("all jobs from that board (no department filter).")

    report = "\n".join(lines)
    out_path = "department_discovery_report.txt"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(report)

    print(f"\nDone. Report written to {out_path}")
    print("Copy this file into your Cowork folder, or just read it directly -")
    print("it's the same format the pipeline's own batch commands produce.")


if __name__ == "__main__":
    main()
