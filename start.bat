@echo off
title Invoice Extractor
cd /d "%~dp0"
echo.
echo  [Invoice Extractor] Avvio in corso...
echo.
call .venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo.
    echo  [ERRORE] Ambiente virtuale non trovato in .venv\
    pause
    exit /b 1
)
echo.
echo  =============================================
echo    INVOICE EXTRACTOR - B2B PDF to Excel
echo  =============================================
echo.
echo  CLI:      invoice-extract ./invoices/
echo  Web UI:   http://localhost:7860
echo.
echo  Scegli:
echo    1 - Web UI (Gradio, browser)
echo    2 - CLI (help menu)
echo    3 - Test (smoke test)
echo.
set /p choice="[1/2/3]: "
if "%choice%"=="1" (
    echo.
    echo  Avvio Web UI su http://localhost:7860...
    python -m invoice_extractor.web
) else if "%choice%"=="2" (
    invoice-extract --help
    echo.
    pause
) else if "%choice%"=="3" (
    python test_smoke.py
    echo.
    pause
) else (
    echo  Scelta non valida. Avvio Web UI...
    python -m invoice_extractor.web
)
pause
