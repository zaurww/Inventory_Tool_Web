@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" goto setup_needed
echo Save and CLOSE Inventory_Input.xlsx and Inventory_Report.xlsx before recalculating.
echo Missing row codes will be assigned automatically. A backup is saved before input changes.
".venv\Scripts\python.exe" "inventory.py" --prepare-input --input "Inventory_Input.xlsx" --output "Inventory_Report.xlsx"
if errorlevel 1 goto failed
start "" "%~dp0Inventory_Report.xlsx"
pause
exit /b 0
:setup_needed
echo Please run Install.bat first.
pause
exit /b 1
:failed
echo.
echo Calculation failed. The previous report was NOT refreshed.
echo Read the error above and Inventory_Report_last_run.txt.
echo For data errors, open Inventory_Report_Checks.xlsx.
pause
exit /b 1
