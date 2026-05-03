@echo off
chcp 65001 >nul
title Supply Stoplight v5 — Updating...
cd /d "%~dp0"

echo.
echo  ╔══════════════════════════════════════╗
echo  ║        supply.stoplight v5           ║
echo  ╚══════════════════════════════════════╝
echo.

:: ── Read default date from settings.py ────────────────────────────────────────
for /f %%i in ('python -c "import sys; sys.path.insert(0,'scripts'); from settings import FECHA_SALDO; print(FECHA_SALDO)"') do set FECHA_DEFAULT=%%i

echo  Current date in settings:  %FECHA_DEFAULT%
echo.
set /p USAR_DEFAULT= Use this date? (Y/N):

if /i "%USAR_DEFAULT%"=="S" goto RUN_DEFAULT
if /i "%USAR_DEFAULT%"=="Y" goto RUN_DEFAULT

:: ── Ask for a custom date ──────────────────────────────────────────────────────
echo.
set /p FECHA_INPUT= Enter date (YYYY-MM-DD):

:: Basic format validation: check length is 10
if not "%FECHA_INPUT:~9,1%"=="" goto FECHA_OK
if "%FECHA_INPUT:~9,1%"=="" if "%FECHA_INPUT:~0,1%"=="" goto FECHA_INVALID
:FECHA_OK
if "%FECHA_INPUT%"=="" goto FECHA_INVALID

echo.
echo  Using entered date: %FECHA_INPUT%
echo.
python scripts\generar_tablero.py --fecha %FECHA_INPUT%
goto CHECK_RESULT

:FECHA_INVALID
echo.
echo  Invalid date. Using default (%FECHA_DEFAULT%).
echo.

:RUN_DEFAULT
echo.
echo  Using default date: %FECHA_DEFAULT%
echo.
python scripts\generar_tablero.py

:CHECK_RESULT
echo.
if %ERRORLEVEL% EQU 0 (
    echo  Dashboard generated successfully.
    echo.
    echo  Opening supply_stoplight.html...
    start "" "supply_stoplight.html"
) else (
    echo  Error generating dashboard.
    echo  Check that Python is installed and Excel files are closed.
)
echo.
pause
