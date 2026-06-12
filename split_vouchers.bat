@echo off
REM ---------------------------------------------------------------
REM  split_vouchers.bat
REM  One-click runner for split_vouchers.py on Windows.
REM
REM  Prerequisites (install once):
REM    1. Python 3.9+ from https://www.python.org/downloads/
REM       (tick "Add python.exe to PATH" during install)
REM    2. Tesseract OCR from https://github.com/UB-Mannheim/tesseract/wiki
REM       (default install path: C:\Program Files\Tesseract-OCR\)
REM    3. Ghostscript from https://www.ghostscript.com/releases/gsdnld.html
REM       (default install path: C:\Program Files\gs\<version>\bin\)
REM    4. Python packages:  py -m pip install pymupdf Pillow
REM ---------------------------------------------------------------

setlocal
cd /d "%~dp0"

REM  --folder "%~dp0." tells the script to scan the directory this .bat
REM  lives in, so users can move/rename the automation folder freely.
REM  Any extra args passed to the .bat (e.g. specific PDFs dragged onto it)
REM  are forwarded as %* and override the folder scan.
where py >nul 2>&1
if %ERRORLEVEL%==0 (
    py "%~dp0split_vouchers.py" --folder "%~dp0." %*
) else (
    python "%~dp0split_vouchers.py" --folder "%~dp0." %*
)

set RC=%ERRORLEVEL%
echo.
if %RC% NEQ 0 (
    echo [split_vouchers] finished with errors ^(exit code %RC%^).
) else (
    echo [split_vouchers] done.
)
pause
endlocal
