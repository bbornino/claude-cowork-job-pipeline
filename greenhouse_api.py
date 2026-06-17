"""
greenhouse_api.py — all Greenhouse-specific logic, isolated from the rest
of the pipeline.

WHY THIS FILE EXISTS
---------------------
daily_pipeline.py used to have Greenhouse's URL shapes, response parsing,
and truncation-recovery code mixed directly into its own logic. That meant
the only way to add a second ATS platform (Lever, Workday, etc.) was to
duplicate or entangle all of that. This module is the template for that:
one file per platform, each exposing the same small interface, so
daily_pipeline.py never needs to know which platform a board uses — it
just calls fetch_jobs(slug) / fetch_job_detail(slug, job_id) /
fetch_departments(slug) and gets back normalized data.

TRANSPORT STRATEGY
-------------------
Every fetch_* function tries a DIRECT outbound HTTP call first. If that
fails for any reason (network block, timeout, DNS, anything), it falls
back to reading a pre-fetched cache file that Cowork wrote via its own
web_fetch tool + the `extract-cache` CLI command in daily_pipeline.py.

This matters because the "Cowork's bash sandbox can't reach the internet"
finding came from one specific test (requests.get() inside Cowork's bash
tool, which hit a 403 at the egress proxy). That's a different network
path than web_fetch (Cowork's own tool, with its own access and its own
~94KB response cap) or than wherever this module itself runs. Hardcoding
"never try direct" bakes in possibly-stale information. Trying direct
first means:
  - If direct access works in some environment, it gets used automatically,
    with no truncation risk at all (no 94KB cap, no header-stripping, no
    JSON-recovery parsing needed) - this is the actual fix for the
    truncation problem, not a more clever recovery parser.
  - If it doesn't work, the fallback is identical to today's behavior.
  - Every result reports which path was used, so this is visible in the
    run output rather than a silent assumption either way.

Direct attempts use a short timeout (5s) so a hung connection doesn't
stall a run waiting to fall back.
"""

import json
import os
import re
from datetime import datetime
from pathlib import Path

PLATFORM = "greenhouse"
DEFAULT_CACHE_DIR = "/tmp/pipeline_cache"


# ---------------------------------------------------------------------------
# Cache directory (shared with daily_pipeline.py's extract-cache commands)
# ---------------------------------------------------------------------------

def cache_dir(override=None):
    return Path(override or os.environ.get("PIPELINE_CACHE_DIR") or DEFAULT_CACHE_DIR)


# ---------------------------------------------------------------------------
# URL / slug helpers
# ---------------------------------------------------------------------------

def slug_from_board_url(url):
    """Extract the Greenhouse slug from either the human-facing
    job-boards.greenhouse.io/{slug} URL (as stored in company-boards.xlsx)
    or the API's boards-api.greenhouse.io/v1/boards/{slug}/... form."""
    m = re.search(r"greenhouse\.io/([^/?]+)", url or "")
    return m.group(1) if m else None


def jobs_url(slug):
    return f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=false"


def jobs_with_content_url(slug):
    # content=true: Greenhouse only includes the `departments` array
    # (along with `offices` and the full description) when content=true.
    # With content=false, departments is omitted entirely - confirmed by
    # a real run that fetched successfully but got 0 departments back for
    # every company. Used only for department derivation, never for the
    # daily pipeline's normal job pull (which deliberately stays on
    # content=false to avoid the description weight).
    return f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=true"


def job_detail_url(slug, job_id):
    return f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs/{job_id}?content=true"


def departments_url(slug):
    return f"https://boards-api.greenhouse.io/v1/boards/{slug}/departments"


def department_url(slug, dept_id):
    return f"https://boards-api.greenhouse.io/v1/boards/{slug}/departments/{dept_id}"


# ---------------------------------------------------------------------------
# Direct HTTP attempt (tried first by every fetch_* function below)
# ---------------------------------------------------------------------------

def _try_direct(url, timeout=5):
    """Returns (json_data, error). error is None on success. Never raises -
    every failure mode (network block, timeout, bad JSON, non-200) is
    captured and returned as a short error string instead."""
    try:
        import requests
    except ImportError:
        return None, "requests not installed"
    try:
        resp = requests.get(url, timeout=timeout, headers={"User-Agent": "Mozilla/5.0"})
    except Exception as e:
        return None, f"{type(e).__name__}: {e}"[:300]
    if resp.status_code != 200:
        deny = resp.headers.get("x-deny-reason")
        snippet = (deny or resp.text or "")[:200]
        return None, f"HTTP {resp.status_code} - {snippet}"
    try:
        return resp.json(), None
    except ValueError as e:
        return None, f"Invalid JSON: {e}"


# ---------------------------------------------------------------------------
# Raw-text parsing / truncation recovery
# (used on cache files, which may be a Cowork web_fetch dump: header lines
# + possibly-truncated JSON)
# ---------------------------------------------------------------------------

def _strip_to_json_start(text):
    """Drop any header lines (URL / redirect / Content-Type / blank) that
    Cowork's web_fetch prepends to a saved tool-result file, returning from
    the first '{' or '[' onward. Returns "" if no JSON start is found."""
    for i, ch in enumerate(text or ""):
        if ch in "{[":
            return text[i:]
    return ""


def extract_jobs_from_raw(text):
    """Extract job objects from a (possibly truncated, possibly
    header-prefixed) `{"jobs": [...]}` board-list response.
    Returns (jobs_list, truncated_bool)."""
    body = _strip_to_json_start(text)
    if not body:
        return [], False
    try:
        data = json.loads(body)
        return data.get("jobs", []), False
    except json.JSONDecodeError:
        pass
    m = re.search(r'"jobs"\s*:\s*\[', body)
    if not m:
        return [], True
    decoder = json.JSONDecoder()
    pos = m.end()
    jobs = []
    while True:
        while pos < len(body) and body[pos] in " \t\n\r,":
            pos += 1
        if pos >= len(body) or body[pos] != "{":
            break
        try:
            obj, pos = decoder.raw_decode(body, pos)
        except json.JSONDecodeError:
            break
        jobs.append(obj)
    return jobs, True


def extract_posting_from_raw(text):
    """Extract a single job-detail (?content=true) object.
    Returns (job_dict_or_None, status) - status in "ok"/"empty"/"truncated"."""
    body = _strip_to_json_start(text)
    if not body:
        return None, "empty"
    try:
        return json.loads(body), "ok"
    except json.JSONDecodeError:
        pass
    try:
        obj, _ = json.JSONDecoder().raw_decode(body, 0)
        return obj, "ok"
    except json.JSONDecodeError:
        return None, "truncated"


def extract_dept_from_raw(text):
    """Extract jobs from a single-department response (GET .../departments/{id})."""
    body = _strip_to_json_start(text)
    if not body:
        return [], "empty"
    try:
        data = json.loads(body)
        return data.get("jobs", []), None
    except json.JSONDecodeError:
        pass
    jobs, truncated = extract_jobs_from_raw(body)
    return jobs, ("truncated" if truncated else None)


def extract_deptlist_from_raw(text):
    """Extract department list from GET .../departments.

    This endpoint embeds each department's full job array inline - there's
    no "departments only, omit job bodies" option in Greenhouse's public
    API - so large boards reliably exceed any fetch-size cap and the
    response gets cut off mid-job-array. A clean json.loads() then fails
    entirely even though department id/name pairs (which appear *before*
    each department's jobs array) are intact.

    Three-stage recovery:
      1. Clean parse - works when the response fit in one fetch.
      2. raw_decode loop over complete top-level department objects, then
         regex over whatever's left (the one that got cut off mid-way) so a
         board with N clean departments + 1 truncated one doesn't silently
         drop the truncated one.
      3. Regex extraction of bare {"id":N,"name":"...",...} pairs - last
         resort when even the first department object is cut off before
         raw_decode can close it.
    Job counts from stages 2-3 are unreliable - callers should flag this."""
    body = _strip_to_json_start(text)
    if not body:
        return [], "empty"

    try:
        data = json.loads(body)
        return data.get("departments", []), None
    except json.JSONDecodeError:
        pass

    m = re.search(r'"departments"\s*:\s*\[', body)
    depts = []
    if m:
        decoder = json.JSONDecoder()
        pos = m.end()
        while True:
            while pos < len(body) and body[pos] in " \t\n\r,":
                pos += 1
            if pos >= len(body) or body[pos] != "{":
                break
            try:
                obj, pos = decoder.raw_decode(body, pos)
                depts.append(obj)
            except json.JSONDecodeError:
                break
        if depts:
            seen_ids = {d.get("id") for d in depts}
            remainder = body[pos:]
            pairs = re.findall(r'\{\s*"id"\s*:\s*(\d+)\s*,\s*"name"\s*:\s*"((?:[^"\\]|\\.)*)"', remainder)
            for did, name in pairs:
                did_int = int(did)
                if did_int not in seen_ids:
                    seen_ids.add(did_int)
                    name = name.replace('\\"', '"').replace("\\\\", "\\")
                    depts.append({"id": did_int, "name": name, "jobs": []})
            return depts, "truncated - job counts may be incomplete"

    pairs = re.findall(r'\{\s*"id"\s*:\s*(\d+)\s*,\s*"name"\s*:\s*"((?:[^"\\]|\\.)*)"', body)
    if pairs:
        seen_ids, depts = set(), []
        for did, name in pairs:
            if did not in seen_ids:
                seen_ids.add(did)
                name = name.replace('\\"', '"').replace("\\\\", "\\")
                depts.append({"id": int(did), "name": name, "jobs": []})
        return depts, "truncated - job counts unavailable (regex recovery)"

    return [], "truncated - could not recover any department data"


# ---------------------------------------------------------------------------
# Public fetch interface — try direct, fall back to cache
# ---------------------------------------------------------------------------

def fetch_jobs_with_content(slug, override_cache_dir=None):
    """Like fetch_jobs, but content=true - used ONLY for department
    derivation (derive_departments_from_jobs needs the `departments` array,
    which Greenhouse omits entirely when content=false). Never used by the
    daily pipeline's normal job pull, which stays on content=false on
    purpose to avoid the description-field weight on every run.

    Returns (jobs_list, error_or_None, source)."""
    data, err = _try_direct(jobs_with_content_url(slug))
    if err is None:
        return data.get("jobs", []), None, "direct"

    cdir = cache_dir(override_cache_dir)
    path = cdir / "boards" / f"{slug}-with-content.json"
    if not path.exists():
        return [], (
            f"Direct fetch failed ({err}) and no cache file at {path}. "
            f"Cowork must web_fetch {jobs_with_content_url(slug)} and run: "
            f"python3 daily_pipeline.py extract-cache board {slug} <result_file> --cache-dir "
            f"(save the result file with a name you'll reference as slug=file)"
        ), "none"
    try:
        return json.loads(path.read_text()).get("jobs", []), None, "cache"
    except Exception as e:
        return [], f"Failed to read cache {path}: {type(e).__name__}: {e}", "none"


def fetch_jobs(slug, override_cache_dir=None, dept_ids=None):
    """Returns (jobs_list, error_or_None, source) where source is
    "direct" or "cache" - daily_pipeline.py surfaces this so it's visible
    which path actually served the data, not assumed.

    No dept_ids: full board job list.
    With dept_ids: merges jobs from each department ID, deduplicated."""
    if not dept_ids:
        data, err = _try_direct(jobs_url(slug))
        if err is None:
            return data.get("jobs", []), None, "direct"

        cdir = cache_dir(override_cache_dir)
        path = cdir / "boards" / f"{slug}.json"
        if not path.exists():
            return [], (
                f"Direct fetch failed ({err}) and no cache file at {path}. "
                f"Cowork must web_fetch {jobs_url(slug)} and run: "
                f"python3 daily_pipeline.py extract-cache board {slug} <result_file>"
            ), "none"
        try:
            return json.loads(path.read_text()).get("jobs", []), None, "cache"
        except Exception as e:
            return [], f"Failed to read cache {path}: {type(e).__name__}: {e}", "none"

    # Department-filtered: try direct per department, merge, dedup
    all_jobs, errors, seen_ids = [], [], set()
    any_direct = False
    for did in dept_ids:
        data, err = _try_direct(department_url(slug, did))
        if err is None:
            any_direct = True
            jobs = data.get("jobs", [])
        else:
            cdir = cache_dir(override_cache_dir)
            path = cdir / "boards" / f"{slug}-dept-{did}.json"
            if not path.exists():
                errors.append(
                    f"Direct fetch for dept {did} failed ({err}) and no cache. Cowork must web_fetch "
                    f"{department_url(slug, did)} and run: "
                    f"python3 daily_pipeline.py extract-cache dept {slug} {did} <result_file>"
                )
                continue
            try:
                jobs = json.loads(path.read_text()).get("jobs", [])
            except Exception as e:
                errors.append(f"Failed to read {path}: {type(e).__name__}: {e}")
                continue
        for j in jobs:
            jid = str(j.get("id", ""))
            if jid not in seen_ids:
                seen_ids.add(jid)
                all_jobs.append(j)
    err = "; ".join(errors) if errors else None
    return all_jobs, err, ("direct" if any_direct and not errors else "mixed/cache")


def fetch_job_detail(slug, job_id, override_cache_dir=None):
    """Returns (job_dict_or_None, error_or_None, source)."""
    data, err = _try_direct(job_detail_url(slug, job_id))
    if err is None:
        return data, None, "direct"

    path = cache_dir(override_cache_dir) / "postings" / f"{job_id}.json"
    if not path.exists():
        return None, (
            f"Direct fetch failed ({err}) and no cache file at {path} for job {job_id}."
        ), "none"
    try:
        cached = json.loads(path.read_text())
    except Exception as e:
        return None, f"Failed to read cache {path}: {type(e).__name__}: {e}", "none"
    status = cached.get("_status")
    if status in ("empty", "truncated"):
        return None, f"Posting {job_id}: cached response was {status}", "cache"
    return cached, None, "cache"


def fetch_departments(slug, override_cache_dir=None):
    """Returns (departments_list, error_or_None, source)."""
    data, err = _try_direct(departments_url(slug))
    if err is None:
        return data.get("departments", []), None, "direct"

    path = cache_dir(override_cache_dir) / "boards" / f"{slug}-departments.json"
    if not path.exists():
        return [], (
            f"Direct fetch failed ({err}) and no cache file at {path}. Cowork must web_fetch "
            f"{departments_url(slug)} and run extract-cache (or list-departments-batch directly "
            f"on the raw result file)."
        ), "none"
    try:
        depts, parse_err = extract_deptlist_from_raw(path.read_text())
        return depts, parse_err, "cache"
    except Exception as e:
        return [], f"Failed to read cache {path}: {type(e).__name__}: {e}", "none"


# ---------------------------------------------------------------------------
# Normalization: raw Greenhouse job dict -> our internal record shape
# ---------------------------------------------------------------------------

def derive_departments_from_jobs(jobs):
    """Extract a department ID -> {name, job_count} mapping directly from a
    ?content=false job-list response, instead of fetching /departments.

    Both endpoints embed full department objects per job, but /departments
    additionally embeds each job's full content/description (the single
    largest field), inflating the response well past what /jobs?content=false
    needs for the same data. Deriving department names/IDs/counts from the
    job list we're already fetching anyway sidesteps /departments entirely -
    no separate fetch, no separate truncation risk on a heavier endpoint.

    The only difference from /departments: a department with zero current
    job postings won't appear here (since it's derived from jobs, not the
    department registry itself). That's an acceptable tradeoff - a
    zero-job department isn't useful for the Departments filter anyway."""
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



    """Map a raw Greenhouse job dict to daily_pipeline.py's posting record.

    Takes the location/name/employment-type helpers as arguments rather
    than importing daily_pipeline (avoids a circular import — this module
    is meant to be importable standalone, and daily_pipeline imports it).

    Greenhouse's job-list payload (?content=false) includes a `departments`
    array per job - independent of whether department-ID filtering is
    configured on the board, captured either way so it's visible downstream
    (seen-postings, All Listings, Shortlist), not just used as a filter."""
    jid = job.get("id")
    jid = str(jid) if jid is not None else None
    title = job.get("title") or "Not Disclosed"
    url = job.get("absolute_url")
    loc_raw = (job.get("location") or {}).get("name") or ""
    updated_at = job.get("updated_at")
    posted_date_obj = None
    posted_date = "Not Disclosed"
    if updated_at:
        try:
            posted_date_obj = datetime.fromisoformat(str(updated_at).replace("Z", "+00:00")).date()
            posted_date = posted_date_obj
        except Exception:
            pass
    depts = job.get("departments") or []
    dept_name = ", ".join(d.get("name", "") for d in depts if d.get("name")) or "Not Disclosed"
    return {
        "Job Title": title,
        "Company Name": normalize_name_fn(company_name),
        "Source": source,
        "Department": dept_name,
        "Location": loc_raw or "Not Disclosed",
        "Location Type": location_type_fn(loc_raw),
        "Posted Date": posted_date,
        "Pay Range": "Not Disclosed",  # not exposed by ?content=false
        "Employment Type": employment_type_fn(title),
        "URL": url,
        "Job ID": jid,
        "_posted_date_obj": posted_date_obj,
    }
