@echo off
REM ===================================================================
REM  Turn OFF the automatic watcher (removes the scheduled task).
REM  Your PDFs and data are untouched; you can re-enable any time by
REM  double-clicking register_watch_task.bat.
REM ===================================================================
schtasks /Delete /TN "EPAL Watch" /F
echo.
echo The "EPAL Watch" task was removed (if it existed).

REM If this machine was the one holding the single-watcher lock, release it
REM immediately so another PC can take over without waiting out the staleness
REM window. Safe no-op if this machine wasn't the lock holder (or portable
REM Python isn't set up yet).
REM NOTE: deliberately flat, no nested ( ) blocks - %EPAL_PY% inside a block
REM that also SETS it (via the call below) would expand at parse time, i.e.
REM to the OLD value, and silently no-op. See epal-verificacao-2026-08-23
REM for the same class of bug in other .bat files. One statement per line
REM sidesteps it entirely.
if not exist "%~dp0_find_python.bat" goto :skip_release
call "%~dp0_find_python.bat" >nul 2>&1
if not defined EPAL_PY goto :skip_release
"%EPAL_PY%" "%~dp0..\src\automation_of_extraction\watch_all.py" --release-lock >nul 2>&1
:skip_release

echo.
pause
