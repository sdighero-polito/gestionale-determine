@echo off
title Gestione Contratti v7.1 - Demanio Marittimo - Comune di Sanremo
echo.
echo ============================================
echo   Gestione Contratti v7.1
echo   Servizio Demanio Marittimo
echo   Comune di Sanremo
echo ============================================
echo.

REM Verifica Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERRORE: Python non trovato. Installare Python 3.9+
    echo https://www.python.org/downloads/
    pause
    exit /b 1
)

REM Verifica/Installa dipendenze
pip show python-docx >nul 2>&1
if errorlevel 1 (
    echo Installazione dipendenze...
    pip install python-docx openpyxl
)

REM Avvio
cd /d "%~dp0"
echo Avvio in corso...
python app.py

if errorlevel 1 (
    echo.
    echo Si e' verificato un errore. Premere un tasto per chiudere.
    pause
)
