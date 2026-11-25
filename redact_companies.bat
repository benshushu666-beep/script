@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
set "PY=%~dp0.venv-redact\Scripts\python.exe"
set "SCRIPT=%~dp0redact.py"
set "FLAG=--redact-companies"

if not exist "%PY%" goto noenv
if "%~1"=="" goto interactive

"%PY%" "%SCRIPT%" %* %FLAG%
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" echo [exit code %RC%]
goto done

:interactive
echo Material Redaction Tool - redact company names
echo ----------------------------------------
set /p "INPUT=Input file(s) or folder: "
set "INPUT=!INPUT:"=!"
if "!INPUT!"=="" goto noinput
set /p "OUT=Output folder (blank = auto): "
set "OUT=!OUT:"=!"
if "!OUT!"=="" (
    "%PY%" "%SCRIPT%" "!INPUT!" %FLAG%
) else (
    "%PY%" "%SCRIPT%" "!INPUT!" -o "!OUT!" %FLAG%
)
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" echo [exit code %RC%]
goto done

:noenv
echo [ERROR] Virtual environment not found: %PY%
goto done

:noinput
echo No input path given.

:done
echo.
pause
