@echo off
REM ===================================================================
REM  EPAL TUA extractor - WATCH MODE (runs continuously).
REM  Double-click once and leave it running: any TUA PDF you drop into
REM  data\inbox is processed automatically. Close the window to stop.
REM ===================================================================
cd /d "%~dp0"

if exist "..\..\.venv\Scripts\python.exe" (
    "..\..\.venv\Scripts\python.exe" pipeline.py --watch
) else (
    python pipeline.py --watch
)

pause
