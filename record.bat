@echo off
title Playwright Recorder

set PYTHON=%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe

echo Starting recorder with system Chrome...
"%PYTHON%" -m playwright codegen --browser chromium --channel chrome --lang python https://sls.3m.com/contaminants
if errorlevel 1 (
    echo Chrome not found, trying Edge...
    "%PYTHON%" -m playwright codegen --browser chromium --channel msedge --lang python https://sls.3m.com/contaminants
)

pause
