@echo off
title CardScan AI
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set "PY=py") else (set "PY=python")
if not exist ".venv\Scripts\python.exe" (
  echo Creating virtual environment...
  %PY% -m venv .venv
  if errorlevel 1 (echo Python/venv setup failed.&pause&exit /b 1)
)
echo Installing required packages...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if not exist ".env" (
  echo.
  echo .env is missing. Copy .env.example to .env and add your Gemini API key.
  pause
  exit /b 1
)
start "" http://127.0.0.1:5000
".venv\Scripts\python.exe" app.py
pause
