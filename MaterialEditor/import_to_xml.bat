@echo off
cd /d "%~dp0"
python material_xlsx.py import %*
pause
