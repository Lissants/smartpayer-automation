@echo off
echo Starting Smartpayer Automation...
echo Please wait while the application loads.
echo.

python smartpayer_gui.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: The application failed to start.
    echo Check the error message above for details.
    echo.
    pause
)