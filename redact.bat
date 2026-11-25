@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
set "PY=%~dp0.venv-redact\Scripts\python.exe"
set "SCRIPT=%~dp0redact.py"

if not exist "%PY%" goto noenv
if "%~1"=="" goto interactive

"%PY%" "%SCRIPT%" %*
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" echo [exit code %RC%]
goto done

:interactive
echo Material Redaction Tool
echo ----------------------------------------
echo Tip: drag files or folders onto redact.bat to process them directly.
set /p "INPUT=Input file(s) or folder: "
set "INPUT=!INPUT:"=!"
if "!INPUT!"=="" goto noinput
set /p "OUT=Output folder (blank = auto): "
set "OUT=!OUT:"=!"
if "!OUT!"=="" (
    "%PY%" "%SCRIPT%" "!INPUT!"
) else (
    "%PY%" "%SCRIPT%" "!INPUT!" -o "!OUT!"
)
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" echo [exit code %RC%]
goto done

:noenv
echo [ERROR] Virtual environment not found: %PY%
echo Rebuild it with:
echo   python -m venv --system-site-packages .venv-redact
echo   .venv-redact\Scripts\python.exe -m pip install pymupdf rapidocr-onnxruntime
goto done

:noinput
echo No input path given.

:done
echo.
pause
