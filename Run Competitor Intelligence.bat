@echo off
title Competitor Intelligence
setlocal

set INSTALL_DIR=%~dp0

echo.
echo  Running Competitor Intelligence...
echo.
python "%INSTALL_DIR%\competitor_intelligence.py"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  Something went wrong. Screenshot this window and send to the team.
    echo.
    pause
)
