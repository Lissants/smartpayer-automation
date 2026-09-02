@echo off
REM ==========================================================
REM  Smartpayer Success Dashboard launcher
REM ==========================================================
cd /d "%~dp0"

echo Starting Smartpayer Success Dashboard...
echo Please wait while the dashboard loads.
echo.

REM Ensure Streamlit (and pandas/altair, which ship with it) is installed
python -c "import streamlit, pandas, altair" 2>NUL
if %ERRORLEVEL% NEQ 0 (
    echo Required packages not found. Installing...
    python -m pip install --upgrade pip
    python -m pip install streamlit pandas altair
)

python -m streamlit run smartpayer_dashboard.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: The dashboard failed to start.
    echo Check the error message above for details.
    echo.
    pause
)
