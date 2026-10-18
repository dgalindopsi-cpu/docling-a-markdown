@echo off
setlocal

rem ── Lanzador del conversor PDF -> Markdown ──────────────────────────────
rem Se coloca en la MISMA carpeta que docling_pdf2md.py y la carpeta venv.
rem Activa el entorno virtual automáticamente y arranca el script.

cd /d "%~dp0"

if not exist "venv\Scripts\activate.bat" (
    echo.
    echo [ERROR] No se encontro la carpeta venv junto a este archivo .bat
    echo Verifica que iniciar_docling.bat este en:
    echo   %~dp0
    echo junto con la carpeta "venv" y "docling_pdf2md.py"
    echo.
    pause
    exit /b 1
)

call venv\Scripts\activate.bat

echo ============================================================
echo   Conversor PDF -^> Markdown
echo ============================================================
echo.
echo   [1] Modo consola
echo   [2] Interfaz grafica
echo.
choice /c 12 /n /m "Elige 1 o 2: "

if errorlevel 2 (
    python docling_pdf2md.py --gui
) else (
    python docling_pdf2md.py
)

endlocal
