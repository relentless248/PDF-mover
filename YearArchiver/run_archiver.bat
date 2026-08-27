@echo off
chcp 65001 >nul
cd /d "%~dp0"
"../python/python.exe" "year_archiver.py"
pause
