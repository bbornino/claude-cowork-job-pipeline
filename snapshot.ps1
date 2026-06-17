# snapshot.ps1 — manual, deliberate git commit of the pipeline's current state.
#
# WHY THIS EXISTS
# -----------------
# Not an autosave, on purpose. After the FUSE-mount staleness issue and the
# near-miss where Cowork reconstructed a binary file from guessed content,
# the goal is "save deliberately, often" — not "save automatically and hope
# the automation itself doesn't introduce the next bug." Run this yourself,
# whenever you want a checkpoint: after a clean daily run, after department
# discovery, before trying something that might go sideways in Cowork.
#
# WHAT IT DOES
# -------------
#   1. Shows you `git status` BEFORE committing, so you see exactly what
#      changed - never commits blind.
#   2. Refuses to proceed if any of the never-commit personal files (per
#      .gitignore) show up as staged - a last-resort check in case
#      .gitignore itself is ever missing or misconfigured.
#   3. Commits with a timestamped message.
#   4. Shows you the commit that resulted, so you can confirm it's what
#      you expected.
#
# USAGE
# ------
#   .\snapshot.ps1                          (auto-generated message)
#   .\snapshot.ps1 "fixed department IDs"   (custom message, appended to the timestamp)
#
# This does NOT push anywhere - it's a local commit only. Push manually
# (`git push`) whenever you're ready, on your own schedule.

param(
    [string]$Message = ""
)

$ErrorActionPreference = "Stop"

# Personal files that must never be committed - mirrors .gitignore. This is
# a belt-and-suspenders check, not a replacement for .gitignore: if
# .gitignore is ever missing, edited, or not respected for some reason,
# this still catches it before the commit happens.
$NeverCommit = @(
    "candidate-profile.md",
    "settings.md",
    "company-boards.xlsx",
    "company-tracker.xlsx",
    "application-history.xlsx",
    "seen-postings.xlsx",
    "job-listings.xlsx"
)

if (-not (Test-Path ".git")) {
    Write-Host "ERROR: no .git folder here. Run this from inside the repo (after 'git init')." -ForegroundColor Red
    exit 1
}

Write-Host "`n--- git status before committing ---`n" -ForegroundColor Cyan
git status

Write-Host "`n--- checking for personal files accidentally staged ---`n" -ForegroundColor Cyan
$staged = git diff --cached --name-only
$blocked = $staged | Where-Object { $NeverCommit -contains $_ }

if ($blocked) {
    Write-Host "STOPPING - these personal files are staged and must NOT be committed:" -ForegroundColor Red
    $blocked | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    Write-Host "`nUnstage them first: git restore --staged <file>" -ForegroundColor Yellow
    Write-Host "Then check why .gitignore didn't catch this before staging anything else." -ForegroundColor Yellow
    exit 1
}

# Also check the working tree (not just staged) for these files existing
# but un-gitignored, in case someone runs this before ever staging anything.
$tracked = git ls-files
$trackedPersonal = $tracked | Where-Object { $NeverCommit -contains $_ }
if ($trackedPersonal) {
    Write-Host "STOPPING - these personal files are already TRACKED by git (committed previously):" -ForegroundColor Red
    $trackedPersonal | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    Write-Host "`nThis means they got committed at some point before .gitignore covered them." -ForegroundColor Yellow
    Write-Host "Removing them now: git rm --cached <file>  (then commit that removal separately)" -ForegroundColor Yellow
    exit 1
}

git add -A

Write-Host "`n--- what's about to be committed ---`n" -ForegroundColor Cyan
$stagedNow = git diff --cached --name-status
if (-not $stagedNow) {
    Write-Host "Nothing changed - nothing to commit." -ForegroundColor Yellow
    exit 0
}
$stagedNow

# Re-check after the add, since git add -A could have staged a personal
# file if .gitignore somehow doesn't cover it.
$stagedAfterAdd = git diff --cached --name-only
$blockedAfterAdd = $stagedAfterAdd | Where-Object { $NeverCommit -contains $_ }
if ($blockedAfterAdd) {
    Write-Host "`nSTOPPING - 'git add -A' just staged a personal file that should be gitignored:" -ForegroundColor Red
    $blockedAfterAdd | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    git restore --staged $blockedAfterAdd
    Write-Host "`nUnstaged it for you. Check .gitignore - this file should be listed there." -ForegroundColor Yellow
    exit 1
}

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm"
if ($Message) {
    $fullMessage = "$timestamp - $Message"
} else {
    $fullMessage = "$timestamp - snapshot"
}

git commit -m "$fullMessage"

Write-Host "`n--- commit created ---`n" -ForegroundColor Green
git log -1 --stat

Write-Host "`nLocal commit only - nothing pushed. Run 'git push' when you're ready." -ForegroundColor Cyan
