# Wonwoo_Key_Column_Merge_Planner 메인 프로젝트 이관 및 파일 정리 계획안

이 계획은 기존 구버전 CWT 생산계획 프로그램 파일을 정리하고, 고도화된 버전인 `Wonwoo_Key_Column_Merge_Planner`를 루트 디렉토리로 이동시켜 단일 메인 프로젝트 구조로 재구성하는 작업을 다룹니다.

## User Review Required

> [!WARNING]
> 이 작업은 루트 디렉토리에 있는 기존 구버전의 파일들(CWT 생산계획 Python 프로그램, 웹 버전 관련 HTML/JS/CSS 파일 등)을 삭제합니다. 이전에 작업하던 이전 버전의 소스코드가 완전히 제거되므로, 필요시 백업이 완료되었는지 확인 부탁드립니다.

## Proposed Changes

### [Root Directory & Subfolder Cleanup]

`Wonwoo_Key_Column_Merge_Planner` 디렉토리 내의 고도화된 파일을 루트 디렉토리로 이동하고, 기존의 미사용 구버전 파일을 안전하게 정리합니다.

#### [DELETE] [Wonwoo_CWT_Planner_Python.cmd](file:///c:/Users/home/.gemini/antigravity/scratch/원우ENG/Wonwoo_CWT_Planner_Python.cmd)
- 기존 구버전 CWT 생산계획 프로그램 실행 스크립트 삭제.

#### [DELETE] [Wonwoo_CWT_Planner_Python.py](file:///c:/Users/home/.gemini/antigravity/scratch/원우ENG/Wonwoo_CWT_Planner_Python.py)
- 기존 구버전 CWT 생산계획 Python 소스코드 삭제.

#### [DELETE] [app.js](file:///c:/Users/home/.gemini/antigravity/scratch/원우ENG/app.js)
- 기존 구버전 웹 프로그램 JS 파일 삭제.

#### [DELETE] [index.html](file:///c:/Users/home/.gemini/antigravity/scratch/원우ENG/index.html)
- 기존 구버전 웹 프로그램 HTML 파일 삭제.

#### [DELETE] [styles.css](file:///c:/Users/home/.gemini/antigravity/scratch/원우ENG/styles.css)
- 기존 구버전 웹 프로그램 CSS 파일 삭제.

#### [DELETE] [xlsx.full.min.js](file:///c:/Users/home/.gemini/antigravity/scratch/원우ENG/xlsx.full.min.js)
- 기존 구버전 웹 프로그램용 라이브러리 삭제.

#### [MODIFY] [README.md](file:///c:/Users/home/.gemini/antigravity/scratch/원우ENG/README.md)
- 루트의 README.md를 고도화 버전인 `Wonwoo_Key_Column_Merge_Planner/README.md`로 대체합니다.

#### [MODIFY] [build_windows_exe.cmd](file:///c:/Users/home/.gemini/antigravity/scratch/원우ENG/build_windows_exe.cmd)
- PyInstaller 빌드 대상을 고도화된 프로그램인 `Wonwoo_Key_Column_Merge_Planner.py`로 변경하고, 생성될 실행파일 이름을 `원우ENG_서열정보_소요자재_자동_취합_프로그램`으로 수정합니다.

#### [NEW] [Wonwoo_Key_Column_Merge_Planner.py](file:///c:/Users/home/.gemini/antigravity/scratch/원우ENG/Wonwoo_Key_Column_Merge_Planner.py)
- `Wonwoo_Key_Column_Merge_Planner/Wonwoo_Key_Column_Merge_Planner.py` 파일을 루트 디렉토리로 이동시킵니다.

#### [NEW] [Wonwoo_Key_Column_Merge_Planner.cmd](file:///c:/Users/home/.gemini/antigravity/scratch/원우ENG/Wonwoo_Key_Column_Merge_Planner.cmd)
- `Wonwoo_Key_Column_Merge_Planner/Wonwoo_Key_Column_Merge_Planner.cmd` 파일을 루트 디렉토리로 이동시킵니다.

#### [NEW] [원우ENG_Key_Column_Merge_프로그램_계획안.md](file:///c:/Users/home/.gemini/antigravity/scratch/원우ENG/원우ENG_Key_Column_Merge_프로그램_계획안.md)
- `Wonwoo_Key_Column_Merge_Planner/원우ENG_Key_Column_Merge_프로그램_계획안.md` 계획안 파일을 루트 디렉토리로 이동시킵니다.

#### [NEW] [Wonwoo_Key_Column_Merge_Planner_EXE/Wonwoo_Key_Column_Merge_Planner.exe](file:///c:/Users/home/.gemini/antigravity/scratch/원우ENG/Wonwoo_Key_Column_Merge_Planner_EXE/Wonwoo_Key_Column_Merge_Planner.exe)
- 기존 배포용 EXE 디렉토리를 루트 경로의 하위 폴더로 이동시킵니다.

#### [DELETE] [Wonwoo_Key_Column_Merge_Planner](file:///c:/Users/home/.gemini/antigravity/scratch/원우ENG/Wonwoo_Key_Column_Merge_Planner)
- 내용물 이동 완료 후 빈 디렉토리가 된 `Wonwoo_Key_Column_Merge_Planner` 폴더를 통째로 삭제하여 단일 프로젝트 구조로 만듭니다.

---

## Verification Plan

### Manual Verification
1. 이동 및 정리 후 루트 디렉토리가 아래 파일들만 가지도록 올바르게 구성되어 있는지 확인합니다:
   - `Wonwoo_Key_Column_Merge_Planner.py`
   - `Wonwoo_Key_Column_Merge_Planner.cmd`
   - `README.md`
   - `원우ENG_Key_Column_Merge_프로그램_계획안.md`
   - `requirements.txt`
   - `requirements-build.txt`
   - `build_windows_exe.cmd`
   - `.gitignore`
   - `Wonwoo_Key_Column_Merge_Planner_EXE/Wonwoo_Key_Column_Merge_Planner.exe`
2. `Wonwoo_Key_Column_Merge_Planner.cmd` 스크립트를 열어 실행 경로 상 오류가 없는지 최종 검증합니다.
