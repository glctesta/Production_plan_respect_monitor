@echo off
title Production Plan Monitor
cd /d "%~dp0"

echo ============================================
echo  Production Plan Monitor - Setup e Avvio
echo ============================================
echo.

:: Cerca Python in percorsi conosciuti (evita alias Microsoft Store)
set PYTHON=

:: Cerca nelle installazioni standard Python
for /f "delims=" %%i in ('dir /b /o-n "C:\Python*" 2^>nul') do (
    if exist "C:\%%i\python.exe" (
        set "PYTHON=C:\%%i\python.exe"
        goto :found
    )
)

:: Cerca in AppData\Local\Programs\Python (installazione utente)
for /f "delims=" %%i in ('dir /b /o-n "C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python*" 2^>nul') do (
    if exist "C:\Users\%USERNAME%\AppData\Local\Programs\Python\%%i\python.exe" (
        set "PYTHON=C:\Users\%USERNAME%\AppData\Local\Programs\Python\%%i\python.exe"
        goto :found
    )
)

:: Cerca in Program Files
for /f "delims=" %%i in ('dir /b /o-n "C:\Program Files\Python*" 2^>nul') do (
    if exist "C:\Program Files\%%i\python.exe" (
        set "PYTHON=C:\Program Files\%%i\python.exe"
        goto :found
    )
)

:: Cerca in Program Files (x86)
for /f "delims=" %%i in ('dir /b /o-n "C:\Program Files (x86)\Python*" 2^>nul') do (
    if exist "C:\Program Files (x86)\%%i\python.exe" (
        set "PYTHON=C:\Program Files (x86)\%%i\python.exe"
        goto :found
    )
)

:: Ultimo tentativo: py launcher
where py >nul 2>&1
if %ERRORLEVEL% equ 0 (
    set PYTHON=py
    goto :found
)

echo ERRORE: Python non trovato!
echo.
echo Installa Python 3.10+ da https://www.python.org/downloads/
echo Assicurati di selezionare "Add Python to PATH" durante l'installazione.
echo.
pause
exit /b 1

:found
echo Python trovato: %PYTHON%
"%PYTHON%" --version
echo.

:: Installa dipendenze
echo Installazione dipendenze...
"%PYTHON%" -m pip install -r requirements.txt --quiet
if %ERRORLEVEL% neq 0 (
    echo.
    echo ERRORE: Installazione dipendenze fallita!
    echo Prova manualmente: "%PYTHON%" -m pip install -r requirements.txt
    pause
    exit /b 1
)
:: Fix per Python 3.12.0: reinstalla markupsafe senza estensioni C
"%PYTHON%" -m pip install --force-reinstall --no-binary markupsafe markupsafe --quiet 2>nul
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
