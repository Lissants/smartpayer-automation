@echo off
title SmartPayer Letter Generator
cd /d "%~dp0"
python script\smartpayer_letter_gui.py
if errorlevel 1 (
    echo.
    echo  The app exited with an error. See message above.
    pause
)
