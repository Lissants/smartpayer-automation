@echo off
REM ==========================================================
REM  Smartpayer Automation - Streamlit UI launcher
REM ==========================================================
cd /d "%~dp0"

echo Starting Smartpayer Automation (Streamlit UI)...
echo Please wait while the application loads.
echo.

REM Ensure Streamlit is installed
python -c "import streamlit" 2>NUL
if %ERRORLEVEL% NEQ 0 (
    echo Streamlit not found. Installing required packages...
    python -m pip install --upgrade pip
    python -m pip install -r requirements_streamlit.txt
)

REM Launch the app (opens in your default browser)
python -m streamlit run smartpayer_app.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: The application failed to start.
    echo Check the error message above for details.
    echo.
    pause
)
