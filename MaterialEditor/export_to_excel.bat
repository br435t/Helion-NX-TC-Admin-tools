@echo off
cd /d "%~dp0"
python material_xlsx.py export %*
if errorlevel 1 (pause & exit /b 1)
start "" "%~dp0helion_material_library.xlsx"
