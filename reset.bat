@echo off
REM ============================================================
REM  reset.bat  —  Clean stale data for a fresh pipeline run
REM
REM  KEEPS (never touched):
REM    company-tracker.xlsx      (your vetted company verdicts)
REM    application-history.xlsx  (your rejection history)
REM    company-boards.xlsx       (board list)
REM    job-sources.xlsx          (source on/off switches)
REM    job-title-filters.xlsx    (SKIP/PURSUE/CHECK rules)
REM    location-filters.xlsx     (allowed locations)
REM    candidate-profile.md      (your resume)
REM    settings.md               (all pipeline settings)
REM    *.md docs                 (pipeline docs/subagents)
REM    daily_pipeline.py         (the script)
REM    *-template.xlsx           (these files)
REM
REM  RESETS (clears data, keeps headers):
REM    seen-postings.xlsx        <- copy from seen-postings-template.xlsx
REM    job-listings.xlsx         <- copy from job-listings-template.xlsx
REM
REM  DELETES (ephemeral / deprecated):
REM    run-log.xlsx
REM    survivors.json
REM    handoff_companies.json
REM    handoff_fit_assessments.json
REM    phase1_report.json
REM    phase2_report.json
REM
REM  Run this from the folder that contains all the pipeline files.
REM ============================================================

echo.
echo === Pipeline Reset ===
echo.
echo This will erase all job listings and seen-postings data.
echo Company verdicts, rejection history, and configuration are NOT affected.
echo.
set /p CONFIRM="Are you sure you want to delete all your data? (y/n): "
if /i not "%CONFIRM%"=="y" (
    echo.
    echo Cancelled. Nothing was changed.
    exit /b 0
)
echo.

REM --- Reset data files from templates ---
echo Resetting seen-postings.xlsx from template...
copy /Y seen-postings-template.xlsx seen-postings.xlsx
if errorlevel 1 ( echo ERROR: could not copy seen-postings-template.xlsx & goto :error )

echo Resetting job-listings.xlsx from template...
copy /Y job-listings-template.xlsx job-listings.xlsx
if errorlevel 1 ( echo ERROR: could not copy job-listings-template.xlsx & goto :error )

REM --- Delete deprecated / ephemeral files ---
echo.
echo Deleting deprecated and ephemeral files...

if exist run-log.xlsx (
    del /Q run-log.xlsx
    echo   Deleted run-log.xlsx
)
if exist survivors.json (
    del /Q survivors.json
    echo   Deleted survivors.json
)
if exist handoff_companies.json (
    del /Q handoff_companies.json
    echo   Deleted handoff_companies.json
)
if exist handoff_fit_assessments.json (
    del /Q handoff_fit_assessments.json
    echo   Deleted handoff_fit_assessments.json
)
if exist phase1_report.json (
    del /Q phase1_report.json
    echo   Deleted phase1_report.json
)
if exist phase2_report.json (
    del /Q phase2_report.json
    echo   Deleted phase2_report.json
)

echo.
echo Done. Ready for a fresh run.
echo   - seen-postings.xlsx: empty (headers only)
echo   - job-listings.xlsx:  empty (headers only)
echo   - All keeper files untouched.
echo.
goto :eof

:error
echo.
echo Reset did NOT complete cleanly. Check the error above.
exit /b 1


REM --- Reset data files from templates ---
echo Resetting seen-postings.xlsx from template...
copy /Y seen-postings-template.xlsx seen-postings.xlsx
if errorlevel 1 ( echo ERROR: could not copy seen-postings-template.xlsx & goto :error )

echo Resetting job-listings.xlsx from template...
copy /Y job-listings-template.xlsx job-listings.xlsx
if errorlevel 1 ( echo ERROR: could not copy job-listings-template.xlsx & goto :error )

REM --- Delete deprecated / ephemeral files ---
echo.
echo Deleting deprecated and ephemeral files...

if exist run-log.xlsx (
    del /Q run-log.xlsx
    echo   Deleted run-log.xlsx
)
if exist survivors.json (
    del /Q survivors.json
    echo   Deleted survivors.json
)
if exist handoff_companies.json (
    del /Q handoff_companies.json
    echo   Deleted handoff_companies.json
)
if exist handoff_fit_assessments.json (
    del /Q handoff_fit_assessments.json
    echo   Deleted handoff_fit_assessments.json
)
if exist phase1_report.json (
    del /Q phase1_report.json
    echo   Deleted phase1_report.json
)
if exist phase2_report.json (
    del /Q phase2_report.json
    echo   Deleted phase2_report.json
)

echo.
echo Done. Ready for a fresh run.
echo   - seen-postings.xlsx: empty (headers only)
echo   - job-listings.xlsx:  empty (headers only)
echo   - All keeper files untouched.
echo.
goto :eof

:error
echo.
echo Reset did NOT complete cleanly. Check the error above.
exit /b 1
