# 🏢 기업 고정자산 및 비품 관리 시스템 (Asset Management System)

> **IT 장비, 사무용 가구, 전자기기 등 사내 모든 고정자산과 비품을 직관적으로 전산 관리하고 실무용 엑셀 장부로 즉시 연동하는 올인원(All-in-One) 솔루션입니다.**

[![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=flat-square&logo=html5&logoColor=white)](https://developer.mozilla.org/ko/docs/Web/HTML)
[![TailwindCSS](https://img.shields.io/badge/TailwindCSS-06B6D4?style=flat-square&logo=tailwindcss&logoColor=white)](https://tailwindcss.com/)
[![Chart.js](https://img.shields.io/badge/Chart.js-FF6384?style=flat-square&logo=chartdotjs&logoColor=white)](https://www.chartjs.org/)
[![SheetJS](https://img.shields.io/badge/SheetJS(XLSX)-107C41?style=flat-square&logo=microsoftexcel&logoColor=white)](https://sheetjs.com/)
[![Python](https://img.shields.io/badge/Python_3.x-3776AB?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=flat-square)](https://opensource.org/licenses/MIT)

---

## ✨ 핵심 기능 (Key Features)

### 1. 🖥️ 설치 없는 초경량 독립형 웹 대시보드 (`index.html`)
- **제로 셋업(Zero Setup)**: 별도 서버나 DB 설치 없이 파일을 더블클릭하는 것만으로 브라우저(Chrome, Edge 등)에서 즉시 실행됩니다.
- **실시간 통계 KPI 대시보드**: 총 자산 수량, 실물 총 보유량, 총 취득가액, 사용 가동률, 수리/유휴/폐기 건수를 실시간 자동 집계합니다.
- **인터랙티브 시각화 차트**: 
  - 부서별 자산 수량 점유율 (도넛 차트)
  - 카테고리별 자산 취득금액 규모 (막대 차트)
- **강력한 검색 및 복합 필터**:
  - 자산코드, 자산명, 실사용자, 시리얼번호, 설치위치 실시간 통합 검색
  - 분류(IT기기/사무가구/전자제품 등), 부서(개발팀/경영지원팀 등), 상태(사용중/유휴/수리중 등) 다중 필터링
  - 테이블 컬럼별 즉시 정렬 (취득일자, 취득금액, 자산코드 등)
- **스마트 자산 등록 & 수정**:
  - 카테고리와 연도를 기반으로 식별 코드를 원클릭 자동 생성 (`AST-IT-2024-001` 등)
  - 상세 정보(단가, 수량, 관리부서, 실사용자, 위치, S/N, 워런티 메모) 완벽 지원
- **선택 일괄 작업**: 다중 체크박스를 통한 대량 자산 일괄 삭제 기능
- **데이터 영구 보존**: 브라우저 로컬 스토리지(LocalStorage)에 실시간 자동 동기화되어 창을 닫아도 안전하게 유지됩니다.

---

### 2. 📊 완벽한 엑셀(Excel) 양방향 연동
- **엑셀 내보내기 (Export)**:
  - 브라우저 상단의 **[엑셀 내보내기]** 버튼 클릭 시, 현재 등록된 모든 자산 목록을 정돈된 열 너비와 한글 헤더가 포함된 `.xlsx` 파일로 즉시 다운로드합니다.
- **엑셀 불러오기 (Import)**:
  - 기존에 작성된 엑셀/CSV 파일을 업로드하면 테이블 데이터로 자동 파싱되어 즉시 대시보드에 병합 등록됩니다.

---

### 3. 🐍 파이썬 엑셀 자동화 스크립트 (`create_excel.py`)
- Python `openpyxl` 라이브러리를 활용하여 실무 양식에 최적화된 독립 엑셀 문서를 생성합니다.
- **포함 시트**:
  1. `고정자산관리대장`: 셀 배경 서식, 테두리, 수량/통화 포맷(`#,##0"원"`), 취득가액 합계 공식(`=F5*G5`), 합계 행(`SUM`), 드롭다운 유효성 검사 내장
  2. `자산현황_요약통계`: 상태별 건수 집계(`COUNTIF`), 카테고리별 취득금액 집계(`SUMIF`) 수식이 적용된 자동 요약 보고서

---

## 📁 파일 구조 (Directory Structure)

```plaintext
├── index.html              # 인터랙티브 웹 자산관리대장 대시보드 (SPA)
├── create_excel.py         # 고품질 수식/서식 적용 엑셀 템플릿 생성 스크립트
├── 깃허브_업로드.bat       # 윈도우 원클릭 깃허브 동기화 배치 파일
└── README.md               # 프로젝트 상세 설명서
```

---

## 🚀 사용 방법 (Getting Started)

### 1) 웹 대시보드 사용하기 (권장)
1. 리포지토리를 클론하거나 다운로드합니다:
   ```bash
   git clone https://github.com/shwmej79-hub/items.git
   ```
2. **`index.html`** 파일을 더블클릭하여 브라우저에서 실행합니다.
3. 내장된 샘플 데이터(맥북, 모니터, 허먼밀러 의자, 법인차량 등 9건)를 바탕으로 등록, 수정, 필터링 기능을 즉시 체험할 수 있습니다.
4. 상단 **[엑셀 내보내기]**를 눌러 실무 장부로 저장해 활용하세요.

### 2) 파이썬 스크립트로 엑셀 양식 생성하기
```bash
# 필수 패키지 설치
pip install openpyxl

# 엑셀 생성 스크립트 실행
python create_excel.py
```
> 실행 완료 시 동일 경로에 `고정자산관리대장_표준양식.xlsx` 파일이 생성됩니다.

---

## 📋 자산관리대장 항목 명세서

| 컬럼명 | 필수 | 데이터 형태 | 설명 및 예시 |
| :--- | :---: | :---: | :--- |
| **자산관리코드** | O | 문자열 | 고유 식별 번호 (예: `AST-IT-2024-001`, 자동생성 가능) |
| **자산명 (품명)** | O | 문자열 | 모델 및 품목명 (예: `MacBook Pro 16인치 M3`) |
| **분류 (카테고리)** | O | 선택 | `IT기기`, `사무가구`, `전자제품`, `차량운반구`, `기타비품` |
| **취득일자** | O | YYYY-MM-DD | 자산 구매 또는 인수 일자 (예: `2024-01-15`) |
| **수량 / 단가** | O | 숫자 | 보유 수량 및 개당 취득단가 (예: `1개 / 3,490,000원`) |
| **취득가액 합계** | O | 숫자 | 수량 × 단가 자동 계산 (예: `3,490,000원`) |
| **관리부서** | O | 문자열 | 소관 부서 (예: `개발팀`, `경영지원팀`, `디자인팀` 등) |
| **실사용자** | 선택 | 문자열 | 실제 사용자 또는 관리 책임자 (예: `김민수 수석`) |
| **보관 및 설치위치** | 선택 | 문자열 | 실물 위치 (예: `본사 5층 개발실 A-12`, `지하 1층 창고`) |
| **자산상태** | O | 선택 | `사용중`, `유휴(보관)`, `수리중`, `폐기예정`, `폐기` |
| **시리얼번호 (S/N)** | 선택 | 문자열 | 하드웨어 고유 일련번호 (예: `C02G99A0MD6R`) |
| **비고 및 특이사항** | 선택 | 텍스트 | 보증 만료일(워런티), 유지보수 이력, 부속품 정보 |

---

## 💡 실무 운영 팁 (Best Practices)
1. **정기 백업 권장**: 브라우저 캐시 삭제 시 데이터가 소실되지 않도록 정기적으로 **[엑셀 내보내기]**를 진행하여 백업본을 보관하세요.
2. **보고서 인쇄 및 PDF 저장**: 브라우저에서 `Ctrl + P`를 누르면 여백과 스타일이 최적화된 인쇄/PDF 저장 모드를 지원합니다.
3. **GitHub Pages 배포**: 본 저장소의 `Settings > Pages`에서 Branch를 `main`으로 설정하면 사내 구성원 누구나 접속 가능한 웹 서비스 URL로 즉시 무료 호스팅할 수 있습니다.

---

## 📄 라이선스 (License)
본 프로젝트는 [MIT License](LICENSE)에 따라 자유롭게 수정, 배포 및 상업적 이용이 가능합니다.
