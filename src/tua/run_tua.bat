@echo off
REM ===================================================================
REM  EPAL TUA extractor - double-click to run.
REM  1) Put your TUA PDF files in the  data\inbox  folder.
REM  2) Double-click this file.
REM  3) When it finishes, check data\processed and data\review.
REM ===================================================================
cd /d "%~dp0"

if exist "..\..\.venv\Scripts\python.exe" (
    "..\..\.venv\Scripts\python.exe" pipeline.py
) else (
    python pipeline.py
)

echo.
echo Finished. Review the messages above.
pause
