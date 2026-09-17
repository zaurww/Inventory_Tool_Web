@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" goto setup_needed
".venv\Scripts\python.exe" "inventory.py" --input "examples\Inventory_Input_Demo.xlsx" --output "examples\Inventory_Report_Demo.xlsx"
if errorlevel 1 goto failed
start "" "%~dp0examples\Inventory_Report_Demo.xlsx"
pause
exit /b 0
:setup_needed
echo Please run Install.bat first.
pause
exit /b 1
:failed
echo Demo failed. Read the error above. Close the demo report in Excel and retry.
pause
exit /b 1
