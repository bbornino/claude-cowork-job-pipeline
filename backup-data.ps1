# backup-data.ps1 — copies your real, irreplaceable data files into a
# local backup/ folder, with versioned rotation so backing up repeatedly
# doesn't just overwrite the last backup.
#
# WHY THIS EXISTS
# -----------------
# snapshot.ps1 (separate script) commits the CODE/CONFIG to git history.
# This script protects the DATA that's deliberately excluded from git
# (per .gitignore) - your real resume, comp targets, target company list,
# vetting verdicts, application history, and live scrape results. None of
# that belongs in git (it's personal), but "not in git" shouldn't mean
# "only one copy exists anywhere." Run this whenever you want a safety net
# before something risky (a Cowork session, a reset, an experimental
# change) - not on a schedule, just deliberately, the same spirit as
# snapshot.ps1.
#
# WHAT GETS BACKED UP
# ---------------------
# Only the files with no "just regenerate it" fallback:
#   candidate-profile.md, settings.md, company-boards.xlsx,
#   company-tracker.xlsx, application-history.xlsx,
#   seen-postings.xlsx, job-listings.xlsx
#
# Deliberately NOT backed up: __pycache__, pipeline_cache/, and the
# ephemeral *_report.json / survivors.json / handoff_*.json files - all of
# those are recreated automatically by the next run, so backing them up
# would just be noise that makes it harder to find the backups that matter.
#
# VERSIONING SCHEME
# -------------------
# First backup of job-listings.xlsx           -> backup\job-listings.backup.xlsx
# Back up again (one already exists)            -> the OLD one is renamed to
#                                                    backup\job-listings.backup.1.xlsx,
#                                                    and the fresh copy takes the
#                                                    unnumbered "<name>.backup.xlsx" name.
# Back up a third time                          -> .backup.1 bumps to .backup.2,
#                                                    fresh copy is .backup.xlsx again.
# So the UNNUMBERED file is always the most recent backup; higher numbers
# are older. Keeps the last 5 versions per file by default (oldest beyond
# that get deleted) - change $KeepVersions below if you want more/fewer.
#
# USAGE
# ------
#   .\backup-data.ps1
#
# The backup\ folder itself should be in .gitignore (this script will add
# it for you if it isn't already there).

$ErrorActionPreference = "Stop"

$FilesToBackup = @(
    "candidate-profile.md",
    "settings.md",
    "company-boards.xlsx",
    "company-tracker.xlsx",
    "application-history.xlsx",
    "seen-postings.xlsx",
    "job-listings.xlsx"
)

$BackupDir = "backup"
$KeepVersions = 5  # how many rotated versions to keep per file, beyond the unnumbered latest

if (-not (Test-Path $BackupDir)) {
    New-Item -ItemType Directory -Path $BackupDir | Out-Null
    Write-Host "Created $BackupDir\" -ForegroundColor Cyan
}

# Make sure backup/ is gitignored - add it if it's missing, rather than
# assuming whoever set up .gitignore remembered this folder.
$gitignorePath = ".gitignore"
$backupIgnoreLine = "backup/"
if (Test-Path $gitignorePath) {
    $gitignoreContent = Get-Content $gitignorePath -Raw
    if ($gitignoreContent -notmatch [regex]::Escape($backupIgnoreLine)) {
        Add-Content -Path $gitignorePath -Value "`n# Local data backups - never commit, these duplicate the personal files above.`n$backupIgnoreLine"
        Write-Host "Added '$backupIgnoreLine' to .gitignore (wasn't there yet)." -ForegroundColor Cyan
    }
} else {
    Write-Host "WARNING: no .gitignore found here. Creating one with just '$backupIgnoreLine' in it - merge with your real .gitignore if this is the wrong folder." -ForegroundColor Yellow
    Set-Content -Path $gitignorePath -Value $backupIgnoreLine
}

Write-Host "`n--- backing up ---`n" -ForegroundColor Cyan

$backedUp = 0
$skipped = 0

foreach ($file in $FilesToBackup) {
    if (-not (Test-Path $file)) {
        Write-Host "  skip (not found): $file" -ForegroundColor DarkGray
        $skipped++
        continue
    }

    $base = [System.IO.Path]::GetFileNameWithoutExtension($file)
    $ext = [System.IO.Path]::GetExtension($file)
    $latestBackupPath = Join-Path $BackupDir "$base.backup$ext"

    if (Test-Path $latestBackupPath) {
        # Rotate existing numbered versions UP first (highest number moves
        # first, so nothing gets overwritten mid-rotation), then move the
        # current unnumbered ".backup" file to ".backup.1".
        for ($i = $KeepVersions; $i -ge 1; $i--) {
            $thisVersioned = Join-Path $BackupDir "$base.backup.$i$ext"
            if (Test-Path $thisVersioned) {
                if ($i -eq $KeepVersions) {
                    Remove-Item $thisVersioned -Force
                    Write-Host "  pruned oldest version: $base.backup.$i$ext (kept last $KeepVersions)" -ForegroundColor DarkGray
                } else {
                    $nextVersioned = Join-Path $BackupDir "$base.backup.$($i + 1)$ext"
                    Move-Item $thisVersioned $nextVersioned -Force
                }
            }
        }
        $firstVersioned = Join-Path $BackupDir "$base.backup.1$ext"
        Move-Item $latestBackupPath $firstVersioned -Force
        Write-Host "  rotated: $base.backup$ext -> $base.backup.1$ext" -ForegroundColor DarkGray
    }

    Copy-Item $file $latestBackupPath -Force
    Write-Host "  saved: $file -> $BackupDir\$base.backup$ext" -ForegroundColor Green
    $backedUp++
}

Write-Host "`n--- done ---" -ForegroundColor Cyan
Write-Host "$backedUp file(s) backed up, $skipped skipped (not found - fine if you haven't created them yet)."
Write-Host "Backups live in .\$BackupDir\ and are gitignored - they're a local safety net, not a git history."
