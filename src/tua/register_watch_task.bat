@echo off
REM ===================================================================
REM  Set up FULLY AUTOMATIC processing with NO terminal window.
REM  Double-click this ONCE. It registers a Windows scheduled task that
REM  checks data\inbox every 5 minutes, in the background, hidden, and
REM  keeps working after the PC restarts.
REM  To turn it off later, double-click unregister_watch_task.bat.
REM ===================================================================
setlocal

REM Prefer the project's silent Python (pythonw = no console window)
for %%I in ("%~dp0..\..\.venv\Scripts\pythonw.exe") do set "PYW=%%~fI"
if not exist "%PYW%" set "PYW=pythonw.exe"
set "SCRIPT=%~dp0pipeline.py"

schtasks /Create /TN "EPAL TUA Watch" /TR "\"%PYW%\" \"%SCRIPT%\"" /SC MINUTE /MO 5 /F

if %errorlevel%==0 (
    echo.
    echo SUCCESS. The inbox is now processed automatically every 5 minutes,
    echo in the background, with NO terminal window - and after every reboot.
    echo Drop TUA PDFs into  data\inbox  and check  data\processed / data\review.
    echo To stop it later, run  unregister_watch_task.bat
) else (
    echo.
    echo Could not create the task. Right-click this file and "Run as administrator".
)
echo.
pause
endlocal
