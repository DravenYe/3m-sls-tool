@echo off
chcp 65001 >nul
title 3M Cartridge Service Life Tool

echo.
echo  Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  [ERROR] Python not found.
    echo  Download: https://www.python.org/downloads/
    echo  During install, check "Add Python to PATH"!
    echo.
    pause
    exit /b 1
)

echo  Python OK. Starting tool...
echo.
python 3m_sls_v2.py
pause
