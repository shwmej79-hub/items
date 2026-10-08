"""
고정자산 및 비품 관리대장 Excel 생성기
사용 라이브러리: openpyxl (pip install openpyxl)
실행 방법: python create_excel.py
결과: 고정자산관리대장_표준양식.xlsx 파일이 생성됩니다.
"""

import sys
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:
    print("[안내] openpyxl 라이브러리가 필요합니다.")
    print("설치 명령어: pip install openpyxl")
    sys.exit(1)

def create_asset_workbook():
    wb = Workbook()
    
    # ----------------------------------------------------
    # 시트 1: 자산관리대장
    # ----------------------------------------------------
    ws = wb.active
    ws.title = "고정자산관리대장"
    ws.views.sheetView[0].showGridLines = True

    # 스타일 정의
    font_title = Font(name="맑은 고딕", size=16, bold=True, color="1E293B")
    font_subtitle = Font(name="맑은 고딕", size=9, color="64748B")
    font_header = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="맑은 고딕", size=10, color="0F172A")
    font_bold = Font(name="맑은 고딕", size=10, bold=True, color="0F172A")
    font_total = Font(name="맑은 고딕", size=11, bold=True, color="1E3A8A")

    fill_header = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid") # 모던 블루
    fill_total = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")  # 연블루
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")  # 연회색교차

    border_thin = Border(
        left=Side(style='thin', color="E2E8F0"),
        right=Side(style='thin', color="E2E8F0"),
        top=Side(style='thin', color="E2E8F0"),
        bottom=Side(style='thin', color="E2E8F0")
    )
    border_total = Border(
        left=Side(style='thin', color="93C5FD"),
        right=Side(style='thin', color="93C5FD"),
        top=Side(style='medium', color="3B82F6"),
        bottom=Side(style='double', color="3B82F6")
    )

    # 상단 타이틀
    ws.merge_cells("A1:M1")
    ws["A1"] = "기업 고정자산 및 비품 관리대장"
    ws["A1"].font = font_title
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 35

    ws.merge_cells("A2:M2")
    ws["A2"] = "※ 본 대장은 기업 내 IT기기, 사무가구, 전자기기 등 모든 비품 및 자산의 이력을 추적 관리하기 위한 표준 장부입니다."
    ws["A2"].font = font_subtitle
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 18

    # 헤더 컬럼
    headers = [
        "연번", "자산관리코드", "자산명 (품명)", "분류(카테고리)", "취득일자", 
        "수량", "취득단가", "취득가액 합계", "관리부서", "실사용자", 
        "보관위치", "자산상태", "시리얼번호 (S/N)", "비고 및 특이사항"
    ]
    
    ws.row_dimensions[4].height = 28
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin

    # 초기 샘플 데이터
    sample_rows = [
        [1, "AST-IT-2024-001", "MacBook Pro 16인치 M3", "IT기기", "2024-01-15", 1, 3490000, "=F5*G5", "개발팀", "김민수 수석", "본사 5층 개발실 A-12", "사용중", "C02G99A0MD6R", "애플케어 플러스 적용 (~2027)"],
        [2, "AST-IT-2024-002", "Dell UltraSharp 32인치 4K", "IT기기", "2024-01-20", 2, 950000, "=F6*G6", "디자인팀", "이서연 팀장", "본사 4층 디자인룸", "사용중", "CN-0R289D-74445", "듀얼 모니터 구성"],
        [3, "AST-FUR-2023-015", "허먼밀러 에어론 체어 B사이즈", "사무가구", "2023-08-10", 5, 1850000, "=F7*G7", "경영지원팀", "임원실/총무팀", "본사 6층 임원실", "사용중", "HM-AER-202308", "12년 본사 워런티 보증"],
        [4, "AST-IT-2023-088", "LG 그램 15인치 노트북", "IT기기", "2023-05-12", 1, 1650000, "=F8*G8", "영업팀", "박준혁 과장", "외근 및 출장용", "수리중", "LG15Z90R-GA56K", "액정 패널 AS센터 접수"],
        [5, "AST-ELC-2023-004", "삼성 비스포크 공기청정기", "전자제품", "2023-03-02", 3, 680000, "=F9*G9", "공용/총무", "사옥 공용", "대회의실 및 라운지", "사용중", "AX53A9310GGD", "필터 정기교체(매년 4월)"],
        [6, "AST-FUR-2022-042", "데스커 전동 모션데스크 1600", "사무가구", "2022-11-05", 1, 520000, "=F10*G10", "개발팀", "미배정", "지하 1층 창고", "유휴(보관)", "DSK-MD-1600W", "신규 입사자 배정 대기"],
        [7, "AST-CAR-2022-001", "현대 아이오닉5 (법인차량)", "차량운반구", "2022-06-18", 1, 54000000, "=F11*G11", "경영지원팀", "법인 공용", "지하 2층 주차장", "사용중", "123가 4567", "업무운행일지 작성 필수"],
        [8, "AST-IT-2021-030", "HP 고속 복합기 MFP-E82660", "IT기기", "2021-09-01", 1, 3200000, "=F12*G12", "공용/총무", "전사 공용", "본사 4층 복사실", "사용중", "HPE82660-KR003", "정기 렌탈 유지보수"],
        [9, "AST-IT-2020-011", "구형 데스크톱 본체 (i5)", "IT기기", "2020-04-10", 2, 850000, "=F13*G13", "개발팀", "불용자산", "지하 창고 폐기구역", "폐기예정", "SYS-OLD-2020", "디가우징 후 폐기 진행"],
    ]

    start_row = 5
    for r_idx, row_data in enumerate(sample_rows, start=start_row):
        ws.row_dimensions[r_idx].height = 22
        is_even = (r_idx % 2 == 0)
        current_fill = fill_zebra if is_even else PatternFill(fill_type=None)

        for c_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.border = border_thin
            if current_fill.fill_type:
                cell.fill = current_fill

            # 서식 적용
            if c_idx in [1, 5, 12]: # 연번, 취득일자, 상태
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx in [2, 13]: # 코드, S/N
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx in [6]: # 수량
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c_idx in [7, 8]: # 단가, 합계금액
                cell.number_format = '#,##0"원"'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # 합계(Summary) 행 추가
    tot_row = start_row + len(sample_rows)
    ws.row_dimensions[tot_row].height = 26
    ws.merge_cells(f"A{tot_row}:E{tot_row}")
    ws[f"A{tot_row}"] = "합 계 (TOTAL)"
    ws[f"A{tot_row}"].font = font_total
    ws[f"A{tot_row}"].alignment = Alignment(horizontal="center", vertical="center")
    
    for c in range(1, 6):
        ws.cell(row=tot_row, column=c).fill = fill_total
        ws.cell(row=tot_row, column=c).border = border_total

    # 수량 합계
    cell_tot_qty = ws.cell(row=tot_row, column=6, value=f"=SUM(F{start_row}:F{tot_row-1})")
    cell_tot_qty.font = font_total
    cell_tot_qty.fill = fill_total
    cell_tot_qty.border = border_total
    cell_tot_qty.number_format = '#,##0'
    cell_tot_qty.alignment = Alignment(horizontal="right", vertical="center")

    # 단가 열 빈칸 처리
    cell_empty = ws.cell(row=tot_row, column=7, value="-")
    cell_empty.font = font_bold
    cell_empty.fill = fill_total
    cell_empty.border = border_total
    cell_empty.alignment = Alignment(horizontal="center", vertical="center")

    # 금액 합계
    cell_tot_val = ws.cell(row=tot_row, column=8, value=f"=SUM(H{start_row}:H{tot_row-1})")
    cell_tot_val.font = font_total
    cell_tot_val.fill = fill_total
    cell_tot_val.border = border_total
    cell_tot_val.number_format = '#,##0"원"'
    cell_tot_val.alignment = Alignment(horizontal="right", vertical="center")

    # 나머지 컬럼 빈칸 테두리
    for c in range(9, 15):
        c_cell = ws.cell(row=tot_row, column=c, value="")
        c_cell.fill = fill_total
        c_cell.border = border_total

    # 열 너비 설정
    col_widths = {
        'A': 7,   # 연번
        'B': 18,  # 자산코드
        'C': 28,  # 자산명
        'D': 14,  # 분류
        'E': 13,  # 취득일자
        'F': 9,   # 수량
        'G': 16,  # 취득단가
        'H': 18,  # 취득가액합계
        'I': 14,  # 관리부서
        'I': 14,  # 실사용자
        'K': 24,  # 보관위치
        'L': 13,  # 자산상태
        'M': 20,  # 시리얼
        'N': 30   # 비고
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    # 드롭다운 유효성 검사 (자산상태, 카테고리)
    dv_status = DataValidation(type="list", formula1='"사용중,유휴(보관),수리중,폐기예정,폐기"', allow_blank=True)
    ws.add_data_validation(dv_status)
    dv_status.add(f"L{start_row}:L100")

    dv_category = DataValidation(type="list", formula1='"IT기기,사무가구,전자제품,차량운반구,기타비품"', allow_blank=True)
    ws.add_data_validation(dv_category)
    dv_category.add(f"D{start_row}:D100")

    # ----------------------------------------------------
    # 시트 2: 대시보드 요약 통계
    # ----------------------------------------------------
    ws_summary = wb.create_sheet(title="자산현황_요약통계")
    ws_summary.views.sheetView[0].showGridLines = True

    ws_summary.merge_cells("A1:E1")
    ws_summary["A1"] = "자산 현황 통계 요약표"
    ws_summary["A1"].font = font_title
    ws_summary["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws_summary.row_dimensions[1].height = 35

    # 상태별 요약 테이블
    ws_summary.row_dimensions[3].height = 24
    ws_summary["A3"] = "상태 구분"
    ws_summary["B3"] = "품목 건수"
    for col in ["A", "B"]:
        ws_summary[f"{col}3"].font = font_header
        ws_summary[f"{col}3"].fill = fill_header
        ws_summary[f"{col}3"].alignment = Alignment(horizontal="center", vertical="center")
        ws_summary[f"{col}3"].border = border_thin

    statuses = ["사용중", "유휴(보관)", "수리중", "폐기예정", "폐기"]
    for idx, st in enumerate(statuses, start=4):
        ws_summary[f"A{idx}"] = st
        ws_summary[f"B{idx}"] = f'=COUNTIF(고정자산관리대장!$L$5:$L$50, "{st}")'
        ws_summary[f"A{idx}"].font = font_data
        ws_summary[f"B{idx}"].font = font_bold
        ws_summary[f"A{idx}"].border = border_thin
        ws_summary[f"B{idx}"].border = border_thin
        ws_summary[f"A{idx}"].alignment = Alignment(horizontal="center", vertical="center")
        ws_summary[f"B{idx}"].alignment = Alignment(horizontal="right", vertical="center")
        ws_summary[f"B{idx}"].number_format = '#,##0"건"'

    # 카테고리별 요약 테이블
    ws_summary["D3"] = "분류(카테고리)"
    ws_summary["E3"] = "취득가액 합계"
    for col in ["D", "E"]:
        ws_summary[f"{col}3"].font = font_header
        ws_summary[f"{col}3"].fill = PatternFill(start_color="4F46E5", end_color="4F46E5", fill_type="solid")
        ws_summary[f"{col}3"].alignment = Alignment(horizontal="center", vertical="center")
        ws_summary[f"{col}3"].border = border_thin

    cats = ["IT기기", "사무가구", "전자제품", "차량운반구", "기타비품"]
    for idx, c in enumerate(cats, start=4):
        ws_summary[f"D{idx}"] = c
        ws_summary[f"E{idx}"] = f'=SUMIF(고정자산관리대장!$D$5:$D$50, "{c}", 고정자산관리대장!$H$5:$H$50)'
        ws_summary[f"D{idx}"].font = font_data
        ws_summary[f"E{idx}"].font = font_bold
        ws_summary[f"D{idx}"].border = border_thin
        ws_summary[f"E{idx}"].border = border_thin
        ws_summary[f"D{idx}"].alignment = Alignment(horizontal="center", vertical="center")
        ws_summary[f"E{idx}"].alignment = Alignment(horizontal="right", vertical="center")
        ws_summary[f"E{idx}"].number_format = '#,##0"원"'

    ws_summary.column_dimensions['A'].width = 16
    ws_summary.column_dimensions['B'].width = 16
    ws_summary.column_dimensions['C'].width = 6
    ws_summary.column_dimensions['D'].width = 18
    ws_summary.column_dimensions['E'].width = 22

    # 저장
    filename = "고정자산관리대장_표준양식.xlsx"
    wb.save(filename)
    print(f"✅ '{filename}' 파일이 성공적으로 생성되었습니다!")
    return filename

if __name__ == "__main__":
    create_asset_workbook()
