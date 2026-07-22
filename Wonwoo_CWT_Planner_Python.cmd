@echo off
setlocal
set "APPDIR=%~dp0"
set "CODEX_PY=C:\Users\HHI\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
cd /d "%APPDIR%"

if exist "%CODEX_PY%" (
  "%CODEX_PY%" "%APPDIR%Wonwoo_CWT_Planner_Python.py"
) else (
  py -3 "%APPDIR%Wonwoo_CWT_Planner_Python.py"
  if errorlevel 1 (
    python "%APPDIR%Wonwoo_CWT_Planner_Python.py"
  )
)

pause
