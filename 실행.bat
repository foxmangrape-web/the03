@echo off
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" goto run
py -3 -c "import sys" >nul 2>&1
if not errorlevel 1 (
    py -3 -m venv .venv
) else (
    python -m venv .venv
)
if errorlevel 1 goto failed
:run
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto failed
echo Open http://127.0.0.1:5000 in your browser.
.venv\Scripts\python.exe app.py
pause
exit /b
:failed
echo Python 3.12 is required. Install Python and try again.
pause
