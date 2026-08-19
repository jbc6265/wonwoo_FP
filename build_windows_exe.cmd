@echo off
setlocal
set "APPDIR=%~dp0"
cd /d "%APPDIR%"

py -3 -m pip install -r requirements-build.txt
py -3 -m PyInstaller ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --name "원우ENG_서열정보_소요자재_자동_취합_프로그램" ^
  "Wonwoo_Key_Column_Merge_Planner.py"

echo.
echo Build complete: dist\원우ENG_서열정보_소요자재_자동_취합_프로그램.exe
pause
