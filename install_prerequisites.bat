@echo off
setlocal EnableDelayedExpansion
title SmartPayer Letter Generator - Prerequisites Installer
color 0A

echo.
echo  ============================================================
echo   SmartPayer Automation Pipeline - Prerequisites Installer
echo  ============================================================
echo.

set ERRORS=0
set WARNINGS=0
set INSTALLER_DIR=%~dp0

:: ── Check admin rights (needed for system-wide installs) ─────────────────────
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo  [NOTE] This script may need admin rights to install Python/LibreOffice.
    echo         If installs fail, re-run as Administrator.
    echo.
)

:: ── [1/4] Python ─────────────────────────────────────────────────────────────
echo [1/4] Checking Python installation...
set PYTHON_FOUND=0

:: Check 1: python in PATH
python --version >nul 2>&1 && set PYTHON_FOUND=1

:: Check 2: python3 in PATH
if !PYTHON_FOUND! equ 0 python3 --version >nul 2>&1 && set PYTHON_FOUND=1

:: Check 3: Registry (64-bit)
if !PYTHON_FOUND! equ 0 (
    reg query "HKLM\SOFTWARE\Python\PythonCore" /s >nul 2>&1 && set PYTHON_FOUND=1
)

:: Check 4: Common install paths
if !PYTHON_FOUND! equ 0 (
    if exist "C:\Python314\python.exe" set PYTHON_FOUND=1
    if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" set PYTHON_FOUND=1
)

if !PYTHON_FOUND! equ 1 (
    for /f "tokens=*" %%i in ('python --version 2^>^&1') do set PYVER=%%i
    echo  [OK] %PYVER%
    goto :python_done
)

:: Python not found - try auto-install
echo  [MISSING] Python 3.9+ not found.
if exist "%INSTALLER_DIR%python-3.14.5-amd64.exe" (
    echo  [AUTO-INSTALL] Found bundled installer. Starting installation...
    echo.
    
    :: Run Python installer silently with PATH option
    "%INSTALLER_DIR%python-3.14.5-amd64.exe" /quiet InstallAllUsers=1 PrependPath=1 TargetDir="C:\Python314"
    
    if errorlevel 1 (
        echo  [ERROR] Python installation failed. Exit code: %errorlevel%
        echo          Please install Python 3.9+ manually from https://python.org
        echo          Make sure to tick "Add Python to PATH".
        set /a ERRORS+=1
        goto :after_packages
    )
    
    :: Wait a moment for PATH to update
    timeout /t 5 /nobreak >nul
    
    :: Re-check Python
    python --version >nul 2>&1
    if errorlevel 1 (
        echo  [WARNING] Python installed but not in PATH. Please restart terminal or add to PATH manually.
        set /a WARNINGS+=1
    ) else (
        for /f "tokens=*" %%i in ('python --version 2^>^&1') do set PYVER=%%i
        echo  [OK] %PYVER% installed successfully.
    )
) else (
    echo  [ERROR] Bundled installer not found: %INSTALLER_DIR%python-3.14.5-amd64.exe
    echo          Please download Python 3.9+ from https://python.org
    echo          Make sure to tick "Add Python to PATH".
    set /a ERRORS+=1
    goto :after_packages
)
:python_done

:: ── [2/4] pip ────────────────────────────────────────────────────────────────
echo.
echo [2/4] Checking pip...
pip --version >nul 2>&1
if errorlevel 1 (
    echo  [WARNING] pip not found. Attempting to bootstrap...
    python -m ensurepip --upgrade >nul 2>&1
    pip --version >nul 2>&1
    if errorlevel 1 (
        echo  [ERROR] Could not install pip. Run: python -m ensurepip --upgrade
        set /a ERRORS+=1
        goto :after_packages
    )
    echo  [OK] pip bootstrapped
) else (
    echo  [OK] pip found
)

:: ── [3/4] Python packages ────────────────────────────────────────────────────
echo.
echo [3/4] Installing Python packages...
echo       python-docx  ^(DOCX reading/writing^)
echo       openpyxl     ^(Excel reading^)
echo       lxml         ^(XML processing^)
echo       watchdog     ^(folder watching^)
echo       tkcalendar   ^(Date Picker^)
echo       pywin32      ^(Windows API^)
echo	      comtypes     ^(MS Word API^)
echo.
pip install python-docx openpyxl lxml watchdog tkcalendar pywin32 comtypes --upgrade
if errorlevel 1 (
    echo  [WARNING] One or more packages may have failed. Check output above.
    set /a WARNINGS+=1
) else (
    echo.
    echo  [OK] All Python packages installed
)

:after_packages

:: ── Check required files in script/ subfolder ─────────────────────────────────
echo.
echo Checking required files in script\ folder...
set MISSING_FILES=0

if not exist "script\smartpayer_letter_generator.py" (
    echo  [MISSING] script\smartpayer_letter_generator.py
    set MISSING_FILES=1
)
if not exist "script\smartpayer_letter_gui.py" (
    echo  [MISSING] script\smartpayer_letter_gui.py
    set MISSING_FILES=1
)
if not exist "script\Smart_Payer_Program_Letter_Template.docx" (
    echo  [MISSING] script\Smart_Payer_Program_Letter_Template.docx
    set MISSING_FILES=1
)
if !MISSING_FILES! equ 0 (
    echo  [OK] All required files present in script\
) else (
    echo  [NOTE] Place all files from the package in the script\ subfolder.
    set /a WARNINGS+=1
)

:: ── Create launcher ───────────────────────────────────────────────────────────
echo.
echo Creating run_gui_letter_generator.bat launcher...
(
    echo @echo off
    echo title SmartPayer Letter Generator
    echo cd /d "%%~dp0"
    echo python smartpayer_gui.py
    echo if errorlevel 1 ^(
    echo     echo.
    echo     echo  The app exited with an error. See message above.
    echo     pause
    echo ^)
) > run_gui_letter_generator.bat
echo  [OK] run_gui.bat created

:: ── Summary ───────────────────────────────────────────────────────────────────
echo.
echo  ============================================================
if !ERRORS! equ 0 (
    if !WARNINGS! equ 0 (
        echo   Setup complete! Everything is ready.
    ) else (
        echo   Setup complete with !WARNINGS! warning^(s^). See notes above.
    )
    echo.
    echo   NOTE: All Python scripts are in the script\ folder.
    echo         lxml handles all XML processing directly.
    echo.
    echo   To launch: double-click run_gui.bat
    echo              or run:  python script\smartpayer_gui.py
) else (
    echo   Setup finished with !ERRORS! error^(s^). Fix issues above first.
)
echo  ============================================================
echo.
pause