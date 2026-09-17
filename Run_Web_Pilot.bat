@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Python environment not found. Run Install.bat first.
  pause
  exit /b 1
)

".venv\Scripts\python.exe" build_web.py
if errorlevel 1 (
  echo Failed to prepare web files.
  pause
  exit /b 1
)

echo.
echo Local web pilot: http://127.0.0.1:8765/
echo Close this window or press Ctrl+C to stop it.
start "" "http://127.0.0.1:8765/"
".venv\Scripts\python.exe" -m http.server 8765 --directory dist --bind 127.0.0.1
