@echo off
title Production Plan Monitor
cd /d "%~dp0"

echo ============================================
echo  Production Plan Monitor - Setup e Avvio
echo ============================================
echo.

:: Cerca Python
set PYTHON=
where python >nul 2>&1 && set PYTHON=python
if not defined PYTHON (
    if exist "C:\Users\User\AppData\Local\Programs\Python\Python311\python.exe" (
        set PYTHON=C:\Users\User\AppData\Local\Programs\Python\Python311\python.exe
    )
)
if not defined PYTHON (
    for /f "delims=" %%i in ('dir /b /s "C:\Python*\python.exe" 2^>nul') do set PYTHON=%%i
)
if not defined PYTHON (
    for /f "delims=" %%i in ('dir /b /s "C:\Users\%USERNAME%\AppData\Local\Programs\Python\*\python.exe" 2^>nul') do set PYTHON=%%i
)

if not defined PYTHON (
    echo ERRORE: Python non trovato!
    echo Installa Python 3.10+ da https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Python trovato: %PYTHON%
echo.

:: Installa dipendenze
echo Installazione dipendenze...
"%PYTHON%" -m pip install -r requirements.txt --quiet
if %ERRORLEVEL% neq 0 (
    echo ERRORE: Installazione dipendenze fallita!
    pause
    exit /b 1
)
echo Dipendenze OK.
echo.

:: Avvia applicazione
echo Avvio Production Plan Monitor sulla porta 8085...
echo Dashboard: http://localhost:8085
echo.
echo Premi CTRL+C per arrestare.
echo ============================================
echo.

"%PYTHON%" app.py

pause
