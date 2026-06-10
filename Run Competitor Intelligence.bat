@echo off
title Competitor Intelligence
setlocal

set INSTALL_DIR=%~dp0

echo.
echo  Running Competitor Intelligence...
echo.

where py >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    py "%INSTALL_DIR%\competitor_intelligence.py"
    goto done
)

where python3 >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    python3 "%INSTALL_DIR%\competitor_intelligence.py"
    goto done
)

python "%INSTALL_DIR%\competitor_intelligence.py"

:done
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  Something went wrong. Screenshot this window and send to the team.
    echo.
    pause
)
