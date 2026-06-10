@echo off
title Competitor Intelligence
setlocal

set INSTALL_DIR=%~dp0

echo.
echo  Checking dependencies...
echo.

where py >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set PYTHON=py
    goto install
)

where python3 >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set PYTHON=python3
    goto install
)

set PYTHON=python

:install
%PYTHON% -m pip install -r "%INSTALL_DIR%\requirements.txt" --quiet

echo.
echo  Running Competitor Intelligence...
echo.

%PYTHON% "%INSTALL_DIR%\competitor_intelligence.py"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  Something went wrong. Screenshot this window and send to the team.
    echo.
    pause
)
