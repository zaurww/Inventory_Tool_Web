@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
echo Inventory Tool - first-time setup
echo Python 3.10 or newer is required. Internet is required for this setup only.
set "INVENTORY_PY_COMMAND="
where py >nul 2>nul
if not errorlevel 1 set "INVENTORY_PY_COMMAND=py -3"
if defined INVENTORY_PY_COMMAND goto check_python
where python >nul 2>nul
if not errorlevel 1 set "INVENTORY_PY_COMMAND=python"
if not defined INVENTORY_PY_COMMAND goto no_python
:check_python
%INVENTORY_PY_COMMAND% -c "import sys; raise SystemExit(0 if sys.version_info >= (3,10) else 1)"
if errorlevel 1 goto no_python
if exist ".venv\Scripts\python.exe" goto install_package
%INVENTORY_PY_COMMAND% -m venv ".venv"
if errorlevel 1 goto failed
:install_package
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r "requirements.txt"
if errorlevel 1 goto failed
echo.
echo Setup complete. Run Run_Demo.bat to check the example.
echo For your records: edit Inventory_Input.xlsx, save it, then run Recalculate.bat.
pause
exit /b 0
:no_python
echo Python 3.10 or newer was not found. Install Python or check the Python launcher.
pause
exit /b 1
:failed
echo Setup failed. Check the error above and your internet connection.
pause
exit /b 1
