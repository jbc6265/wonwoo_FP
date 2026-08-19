@echo off
setlocal
set "APPDIR=%~dp0"
set "CODEX_PY=C:\Users\HHI\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
cd /d "%APPDIR%"

if exist "%CODEX_PY%" (
  "%CODEX_PY%" "%APPDIR%Wonwoo_Key_Column_Merge_Planner.py"
) else (
  py -3 "%APPDIR%Wonwoo_Key_Column_Merge_Planner.py"
  if errorlevel 1 (
    python "%APPDIR%Wonwoo_Key_Column_Merge_Planner.py"
  )
)

pause
