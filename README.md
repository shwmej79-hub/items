# 🏢 기업 고정자산 및 비품 관리 시스템 (Asset Management System)

> **IT 장비, 사무용 가구, 전자기기 등 사내 모든 고정자산과 비품을 직관적으로 전산 관리하고, 수퍼베이스(Supabase) 클라우드 DB 및 실무용 엑셀 장부로 즉시 연동하는 올인원(All-in-One) 솔루션입니다.**

[![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=flat-square&logo=html5&logoColor=white)](https://developer.mozilla.org/ko/docs/Web/HTML)
[![TailwindCSS](https://img.shields.io/badge/TailwindCSS-06B6D4?style=flat-square&logo=tailwindcss&logoColor=white)](https://tailwindcss.com/)
[![Supabase](https://img.shields.io/badge/Supabase-3FCF8E?style=flat-square&logo=supabase&logoColor=white)](https://supabase.com/)
[![Chart.js](https://img.shields.io/badge/Chart.js-FF6384?style=flat-square&logo=chartdotjs&logoColor=white)](https://www.chartjs.org/)
[![SheetJS](https://img.shields.io/badge/SheetJS(XLSX)-107C41?style=flat-square&logo=microsoftexcel&logoColor=white)](https://sheetjs.com/)
[![Python](https://img.shields.io/badge/Python_3.x-3776AB?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=flat-square)](https://opensource.org/licenses/MIT)

---

## ✨ 핵심 기능 (Key Features)

### 1. ⚡ 수퍼베이스(Supabase) 클라우드 DB 연동 지원
- **클라우드 PostgreSQL 연동**: 별도 백엔드 서버 없이 브라우저에서 Supabase API와 직접 통신하여 모든 자산 데이터를 실시간 동기화합니다.
- **하이브리드 데이터 아키텍처**:
  - **Supabase 모드**: 프로젝트 URL과 Anon Key만 입력하면 클라우드 DB와 실시간 CRUD 동기화
  - **로컬 모드(Fallback)**: Supabase 미연결 시 브라우저 LocalStorage로 안전하게 자동 동작
- **로컬 ➔ 클라우드 1-Click 마이그레이션**: 로컬에 등록된 기존 자산 데이터를 Supabase DB로 한 번에 업로드할 수 있는 마이그레이션 버튼 제공
- **보안 및 RLS 정책 완비**: 익명(anon) 키로 안전하게 읽기/쓰기가 가능하도록 Row Level Security(RLS) 및 트리거가 구성된 `supabase_schema.sql` 스크립트 기본 제공

---

### 2. 🖥️ 설치 없는 초경량 독립형 웹 대시보드 (`index.html`)
- **제로 셋업(Zero Setup)**: 브라우저(Chrome, Edge 등)로 파일을 열기만 하면 즉시 실행됩니다.
- **실시간 통계 KPI 대시보드**: 총 자산 수량, 실물 보유량, 총 취득가액, 사용 가동률, 수리/유휴/폐기 건수를 자동 집계합니다.
- **인터랙티브 시각화 차트**:
  - 부서별 자산 수량 점유율 (도넛 차트 - Chart.js)
  - 카테고리별 자산 취득금액 규모 (막대 차트 - Chart.js)
- **강력한 검색 및 복합 필터**:
  - 자산코드, 자산명, 실사용자, 시리얼번호, 설치위치 실시간 통합 검색
  - 분류(IT기기/사무가구/전자제품 등), 부서(개발팀/경영지원팀 등), 상태(사용중/유휴/수리중 등) 다중 필터링
  - 테이블 컬럼별 정렬 (취득일자, 취득금액, 자산코드 등)
- **스마트 자산 등록 & 수정**:
  - 카테고리와 연도를 기반으로 식별 코드 원클릭 자동 생성 (`AST-IT-2024-001` 등)
  - 상세 정보(단가, 수량, 관리부서, 실사용자, 위치, S/N, 워런티 메모) 완벽 지원
- **선택 일괄 작업**: 다중 체크박스를 통한 대량 자산 일괄 삭제

---

### 3. 📊 완벽한 엑셀(Excel) 양방향 연동
- **엑셀 내보내기 (Export)**:
  - 브라우저 상단의 **[엑셀 내보내기]** 버튼 클릭 시, 현재 등록된 모든 자산 목록을 정돈된 열 너비와 한글 헤더가 포함된 `.xlsx` 파일로 즉시 다운로드합니다.
- **엑셀 불러오기 (Import)**:
  - 기존에 작성된 엑셀/CSV 파일을 업로드하면 테이블 데이터로 자동 파싱되어 즉시 대시보드 및 Supabase 클라우드에 일괄 저장됩니다.

---

### 4. 🐍 파이썬 엑셀 자동화 스크립트 (`create_excel.py`)
- Python `openpyxl` 라이브러리를 활용하여 실무 양식에 최적화된 독립 엑셀 문서를 생성합니다.
- **포함 시트**:
  1. `고정자산관리대장`: 셀 배경 서식, 테두리, 수량/통화 포맷(`#,##0"원"`), 취득가액 합계 공식(`=F5*G5`), 합계 행(`SUM`), 드롭다운 유효성 검사 내장
  2. `자산현황_요약통계`: 상태별 건수 집계(`COUNTIF`), 카테고리별 취득금액 집계(`SUMIF`) 수식이 적용된 자동 요약 보고서

---

## 📁 파일 구조 (Directory Structure)

```plaintext
├── index.html              # 인터랙티브 웹 자산관리대장 대시보드 (Supabase 연동 포함)
├── supabase_schema.sql     # Supabase 테이블/RLS/트리거 자동 생성 SQL 스크립트
├── create_excel.py         # 고품질 수식/서식 적용 엑셀 템플릿 생성 스크립트
├── 깃허브_업로드.bat       # 윈도우 원클릭 깃허브 동기화 배치 파일
└── README.md               # 프로젝트 상세 설명서
```

---

## ⚡ Supabase 클라우드 DB 연동 가이드

### 1단계: Supabase 프로젝트 생성 및 SQL 스크립트 실행
1. [Supabase](https://supabase.com)에 로그인 후 새 프로젝트를 생성합니다.
2. 좌측 사이드바 메뉴에서 **[SQL Editor]**로 들어간 뒤 **[New query]**를 클릭합니다.
3. 프로젝트 내 **`supabase_schema.sql`** 파일의 내용을 전체 복사하여 붙여넣고 우측 하단 **[Run]**을 누릅니다.  
   *(assets 테이블, 인덱스, RLS 정책, 샘플 데이터가 1초 만에 자동 생성됩니다.)*

### 2단계: API Key 확인
1. 좌측 하단 톱니바퀴 아이콘 **[Project Settings]** -> **[Data API]** (또는 **[API]**) 메뉴로 이동합니다.
2. **Project URL** (예: `https://xxxx.supabase.co`)과 **anon public API key** 값을 복사합니다.

### 3단계: 웹 대시보드에서 연동
1. 브라우저에서 `index.html`을 엽니다.
2. 상단 헤더의 **[DB 연동 설정]** 버튼을 클릭합니다.
3. 복사한 **Project URL**과 **Anon Key**를 입력하고 **[연결 및 저장]**을 누르면 녹색 불(`🟢 Supabase 연동됨`)이 들어오며 클라우드 동기화가 활성화됩니다!
4. *(기존 로컬 데이터가 있다면 모달 내 [클라우드로 올리기]를 눌러 한 번에 업로드할 수 있습니다.)*

---

## 📋 자산관리대장 항목 명세서

| 컬럼명 | DB 컬럼 | 타입 | 설명 및 예시 |
| :--- | :--- | :---: | :--- |
| **자산관리코드** | `code` | TEXT (UNIQUE) | 고유 식별 번호 (예: `AST-IT-2024-001`, 자동생성 가능) |
| **자산명 (품명)** | `name` | TEXT | 모델 및 품목명 (예: `MacBook Pro 16인치 M3`) |
| **분류 (카테고리)** | `category` | TEXT | `IT기기`, `사무가구`, `전자제품`, `차량운반구`, `기타비품` |
| **취득일자** | `date` | DATE | 자산 구매 또는 인수 일자 (예: `2024-01-15`) |
| **수량 / 단가** | `quantity` / `price` | INTEGER / NUMERIC | 보유 수량 및 개당 취득단가 (예: `1개 / 3,490,000원`) |
| **관리부서** | `department` | TEXT | 소관 부서 (예: `개발팀`, `경영지원팀`, `디자인팀` 등) |
| **실사용자** | `user` | TEXT | 실제 사용자 또는 관리 책임자 (예: `김민수 수석`) |
| **보관 및 설치위치** | `location` | TEXT | 실물 위치 (예: `본사 5층 개발실 A-12`, `지하 1층 창고`) |
| **자산상태** | `status` | TEXT | `사용중`, `유휴(보관)`, `수리중`, `폐기예정`, `폐기` |
| **시리얼번호 (S/N)** | `serial` | TEXT | 하드웨어 고유 일련번호 (예: `C02G99A0MD6R`) |
| **비고 및 특이사항** | `remarks` | TEXT | 보증 만료일(워런티), 유지보수 이력, 부속품 정보 |

---

## 📄 라이선스 (License)
본 프로젝트는 [MIT License](LICENSE)에 따라 자유롭게 수정, 배포 및 상업적 이용이 가능합니다.
