@echo off
REM Turn OFF the automatic background processing set up by register_watch_task.bat.
schtasks /Delete /TN "EPAL TUA Watch" /F
if %errorlevel%==0 (
    echo.
    echo Automatic background processing has been turned OFF.
) else (
    echo.
    echo Task not found (it may already be off).
)
echo.
pause
