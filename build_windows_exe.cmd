@echo off
setlocal
set "APPDIR=%~dp0"
cd /d "%APPDIR%"

py -3 -m pip install -r requirements-build.txt
py -3 -m PyInstaller ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --name "원우ENG_CWT_생산계획_프로그램" ^
  "Wonwoo_CWT_Planner_Python.py"

echo.
echo Build complete: dist\원우ENG_CWT_생산계획_프로그램.exe
pause
