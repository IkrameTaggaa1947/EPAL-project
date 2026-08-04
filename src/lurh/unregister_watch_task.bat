@echo off
REM ===================================================================
REM  Turn OFF the automatic background processing.
REM  Double-click this file to remove the "EPAL LURH Watch" scheduled
REM  task created by register_watch_task.bat.
REM ===================================================================

schtasks /Delete /TN "EPAL LURH Watch" /F

if %errorlevel%==0 (
    echo.
    echo Done. Automatic processing is now OFF.
    echo You can still process files manually with run_lurh.bat.
) else (
    echo.
    echo The task was not found ^(it may already be removed^).
)
echo.
pause
