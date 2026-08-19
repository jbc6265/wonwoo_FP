# 원우ENG 생산계획 및 키 컬럼 병합 프로그램 구조 분석

이 문서는 `wonwoo_FP` 저장소의 루트 디렉토리와 `Wonwoo_Key_Column_Merge_Planner` 디렉토리에 포함된 파일 및 코드 구조를 분석한 내용입니다.

---

## 1. 전체 디렉토리 및 파일 구성

저장소는 크게 두 가지 핵심 기능(생산계획 수립, 키 컬럼 병합 플래너)으로 나뉘며, 각각 **Python GUI 버전**과 **웹(HTML/JS) 버전**을 지원하도록 구성되어 있습니다.

```
원우ENG (Repository Root)
├── .gitignore
├── README.md                              # 루트 생산계획 프로그램 안내 및 기준 정의
├── Wonwoo_CWT_Planner_Python.cmd          # Python CWT 생산계획 프로그램 실행 스크립트
├── Wonwoo_CWT_Planner_Python.py           # Python Tkinter 기반 CWT 생산계획 GUI 프로그램
├── app.js                                 # 웹 버전 CWT 생산계획 프로그램의 비즈니스 로직
├── index.html                             # 웹 버전 CWT 생산계획 프로그램의 UI 마크업
├── styles.css                             # 웹 버전 스타일시트
├── xlsx.full.min.js                       # 웹 버전용 SheetJS 라이브러리
├── requirements.txt                       # Python 실행을 위한 라이브러리 목록 (openpyxl)
├── requirements-build.txt                 # EXE 빌드용 라이브러리 목록 (pyinstaller)
├── build_windows_exe.cmd                  # PyInstaller를 사용한 EXE 빌드 스크립트
│
└── Wonwoo_Key_Column_Merge_Planner        # 4개 조립라인 취합 및 병합 프로그램 폴더
    ├── README.md                          # 병합 프로그램 상세 설명 및 예외 처리 규칙
    ├── Wonwoo_Key_Column_Merge_Planner.cmd # 병합 프로그램 실행 스크립트
    ├── Wonwoo_Key_Column_Merge_Planner.py  # Tkinter 기반 서열정보&소요자재 자동 취합 GUI 프로그램
    ├── 원우ENG_Key_Column_Merge_프로그램_계획안.md # 서열정보 병합 프로그램 요구사항 계획안
    └── Wonwoo_Key_Column_Merge_Planner_EXE (빌드 산출물 혹은 배포용 폴더)
```

---

## 2. 코드 및 기능 분석

### A. CWT 생산계획 수립 프로그램 (루트 디렉토리)

이 프로그램은 월확정조립서열계획(통합1라인, 통합2라인)과 자재소요현황을 분석하여 **카운터웨이트(CWT) 생산 계획**을 생성합니다.

#### ① Python GUI 버전 ([`Wonwoo_CWT_Planner_Python.py`](file:///c:/Users/home/.gemini/antigravity/scratch/원우ENG/Wonwoo_CWT_Planner_Python.py))
- **기술 스택:** Python 3, `tkinter` (GUI), `openpyxl` (Excel 읽기/쓰기).
- **아키텍처 및 흐름:**
  1. **UI 구성 (`PlannerGui` 클래스):** `tk.Tk`를 상속하여 구현. ttk 스타일을 적용한 모던한 카드형 레이아웃 제공. 파일 선택(통합1라인, 통합2라인, 자재소요현황, 출력폴더) 및 생산계획 미리보기 테이블(TreeView)을 보여줌.
  2. **헤더 자동 감지 및 파싱 (`detect_header_row`, `read_rows`):** Excel 파일 내의 헤더 행 위치를 키워드 스코어링 방식으로 자동 감지.
  3. **데이터 필터링 및 매칭 규칙 (`process`):**
     - **통합1라인:** 전체 사용. 자재 코드 중 `KPA10~` 또는 `KSA10~` 패턴 매칭.
     - **통합2라인:** 현대 제품(영업모델이 `HX~`, `R~`로 시작하는 항목)만 사용. 디벨론(`DX~`로 시작) 제품은 생산 제외 및 예외목록 기록.
     - **자재소요현황:** 품명에 `카운터웨이트`가 포함된 건만 필터링.
     - **매칭:** `생산번호 = 물류번호` 기준으로 매칭하여 하나의 행(`PLAN_COLUMNS`)으로 조인.
  4. **엑셀 생성 (`write_sheet`):** `생산계획` (정상 매칭 건), `예외목록` (미매칭, 중복, 필터링 제외 건), `검증요약` (처리 건수 및 파일 정보 요약) 시트를 생성. 착수일 최소/최대값 기준으로 파일명 생성.

#### ② 웹 버전 ([`index.html`](file:///c:/Users/home/.gemini/antigravity/scratch/원우ENG/index.html), [`app.js`](file:///c:/Users/home/.gemini/antigravity/scratch/원우ENG/app.js))
- **기술 스택:** Pure HTML5/CSS3, JavaScript (ES6+), `SheetJS (xlsx.full.min.js)`.
- **특징:** Python GUI의 비즈니스 로직과 동일하게 작동하며, 웹 브라우저 내에서 직접 엑셀을 로컬 파싱하고 병합하여 즉시 다운로드할 수 있도록 설계됨. PC에 Python 설치가 불가능한 환경을 대응하기 위한 대안으로 제공됨.

---

### B. 조립서열&소요자재 자동 취합 프로그램 ([`Wonwoo_Key_Column_Merge_Planner`](file:///c:/Users/home/.gemini/antigravity/scratch/원우ENG/Wonwoo_Key_Column_Merge_Planner))

4개 조립 라인(통합1라인, 통합2라인, 선진정공, 초대형)의 월확정 서열 데이터와 물류번호별 자재소요현황을 분석하여 **선택한 임의의 컬럼들을 키 컬럼(`생산번호=물류번호`) 기준으로 동적 병합**하는 고도화된 취합 도구입니다.

#### ① 소스코드 구조 ([`Wonwoo_Key_Column_Merge_Planner.py`](file:///c:/Users/home/.gemini/antigravity/scratch/원우ENG/Wonwoo_Key_Column_Merge_Planner/Wonwoo_Key_Column_Merge_Planner.py))
- **기술 스택:** Python, `tkinter`, `openpyxl`
- **핵심 아키텍처 및 특징:**
  1. **다중 파일 탭 구조 UI:**
     - 4개의 월확정 라인(통합1, 통합2, 선진정공, 초대형)과 1개의 자재소요현황 파일 선택 지원.
     - 파일별로 실제 헤더를 분석하여 체크박스 목록으로 노출. 사용자가 결과 엑셀에 포함시킬 컬럼을 직접 선택 가능 (검색 및 전체 선택/해제 지원).
     - 라이트/다크 테마 지원 및 파일별로 고유 테마 컬러(Tone)를 부여하여 UI 시인성을 극대화함.
  2. **보안/표준 예외 대응 (DRM 대응 및 Excel Interop):**
     - openpyxl로 직접 열리지 않는 엑셀 파일(보안 파일 혹은 비표준 표준 xml 구조)에 대비하여, Windows 환경일 경우 PowerShell ComObject를 사용하여 백그라운드에서 임시로 정상 포맷 복사본(`.xlsx`)을 만들어 파싱하는 `excel_com_convert_to_xlsx` 및 `open_workbook_with_fallback` 로직 탑재.
  3. **세분화된 병합 및 제외 비즈니스 로직 (`process_merge`):**
     - **Key 매핑:** 월확정(`생산번호`) $\leftrightarrow$ 자재소요(`물류번호`)로 자동 통일.
     - **제외 규칙:** 
       - 모든 라인의 `DX~` 영업모델 제외 및 예외목록 기록.
       - 선진정공 라인의 `WX~`, `HW~`, `R~` 모델은 정상 병합하되, 영업모델 명에 `ACR` 또는 `ECR`이 포함된 경우 제외.
       - 초대형 라인의 `HX1000`, `DX1000`, `DX800` 제외.
     - **중복 자재 우선순위:** 동일 물류번호 자재가 여러 개인 경우 `카운터웨이트`, `COUNTER WEIGHT`, `COUNTERWEIGHT` 키워드를 포함하는 자재를 우선순위로 매칭. 매칭 대상이 복수이거나 없으면 예외 시트에 분류.
     - **중복 키 처리:** 월확정 라인 간 생산번호가 중복될 경우, 누락 없이 각각 모든 행을 결과에 보존하고 `월확정 Key 중복` 예외 기록.
  4. **출력물 품질 향상:**
     - 엑셀 결과 시트(`병합결과`, `예외목록`, `검증요약`)에 고정 헤더, 필터링 영역, 최적의 컬럼 폭 자동 계산(east_asian_width 반영) 적용.

---

## 3. 핵심 규칙 비교 요약

| 구분 | CWT 생산계획 프로그램 (루트) | 서열정보&소요자재 자동 취합 프로그램 (Merge Planner) |
| :--- | :--- | :--- |
| **대상 라인** | 통합1라인, 통합2라인 (2개) | 통합1라인, 통합2라인, 선진정공, 초대형 (4개) |
| **대상 컬럼** | 고정된 컬럼 레이아웃 (`PLAN_COLUMNS`) | 사용자가 각 탭에서 체크박스로 동적 선택 |
| **예외 필터** | 통합2라인의 `DX~` (제외), 현대 이외 모델 제외 | 4개 라인 공통 `DX~` 제외<br>선진정공 `ACR`, `ECR` 제외<br>초대형 `HX1000`, `DX1000`, `DX800` 제외 |
| **자재 필터** | KPA10/20, KSA10/20 & 품명 내 `카운터웨이트` 필수 | 물류번호 ↔ 생산번호 매칭 우선. 중복 발생 시 CWT 키워드 우선순위 필터링 |
| **특이 사항** | 웹 버전(HTML/JS) 병행 지원 | 보안(DRM) 파일 파싱 지원을 위해 PowerShell 기반 ComObject 변환 기능 제공 |
