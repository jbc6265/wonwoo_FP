@echo off
setlocal
set "APPDIR=%~dp0"
cd /d "%APPDIR%"

py -3 "%APPDIR%Wonwoo_CWT_Planner_Python.py"
if errorlevel 1 (
  python "%APPDIR%Wonwoo_CWT_Planner_Python.py"
)

pause
