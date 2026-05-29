@echo off
REM SlideCraft Lite launcher (Windows cmd.exe)
setlocal
cd /d "%~dp0"
if not exist .venv (
    echo ^>^> First run: creating .venv
    python -m venv .venv
    if errorlevel 1 (
        echo Failed to create venv. Install Python 3.10+ and add it to PATH.
        exit /b 1
    )
)
call .venv\Scripts\activate.bat
python -c "import flask" >NUL 2>&1
if errorlevel 1 (
    echo ^>^> Installing dependencies (one-time, ~30 s)
    pip install --upgrade pip
    pip install -r requirements.txt
)
where soffice >NUL 2>&1
if errorlevel 1 (
    if not exist "C:\Program Files\LibreOffice\program\soffice.exe" (
        echo !! LibreOffice not found — PPTX conversion will use a low-fidelity fallback.
        echo !! Install from https://libreoffice.org/download
    )
)
python app.py
endlocal
