-- ==============================================================================
-- 🏢 기업 고정자산 및 비품 관리대장 - Supabase 테이블 생성 스키마
-- ==============================================================================
-- Supabase 대시보드 (https://supabase.com) 접속 -> 프로젝트 선택 
-- -> 좌측 메뉴 [SQL Editor] -> [New query]에 아래 내용을 전체 붙여넣고 [Run]을 누르세요.
-- ==============================================================================

-- 1. assets 테이블 생성 (이미 존재하는 경우 보존)
CREATE TABLE IF NOT EXISTS public.assets (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    code TEXT NOT NULL UNIQUE,                       -- 자산관리코드 (예: AST-IT-2024-001)
    name TEXT NOT NULL,                             -- 자산명 (품명)
    category TEXT NOT NULL,                         -- 분류 (IT기기, 사무가구, 전자제품 등)
    date DATE NOT NULL,                             -- 취득일자
    quantity INTEGER NOT NULL DEFAULT 1 CHECK (quantity > 0), -- 수량
    price NUMERIC NOT NULL DEFAULT 0 CHECK (price >= 0),       -- 취득단가 (원)
    department TEXT NOT NULL,                       -- 관리부서
    "user" TEXT DEFAULT '',                         -- 실사용자 / 담당자
    location TEXT DEFAULT '',                       -- 보관 및 설치위치
    status TEXT NOT NULL DEFAULT '사용중',           -- 상태 (사용중, 유휴(보관), 수리중, 폐기예정, 폐기)
    serial TEXT DEFAULT '',                         -- 시리얼번호 / 모델식별자
    remarks TEXT DEFAULT '',                        -- 비고 및 특이사항
    created_at TIMESTAMPTZ DEFAULT NOW(),           -- 등록일시
    updated_at TIMESTAMPTZ DEFAULT NOW()            -- 최종수정일시
);

-- 2. 검색 및 필터 속도 향상을 위한 인덱스 생성
CREATE INDEX IF NOT EXISTS idx_assets_code ON public.assets(code);
CREATE INDEX IF NOT EXISTS idx_assets_category ON public.assets(category);
CREATE INDEX IF NOT EXISTS idx_assets_department ON public.assets(department);
CREATE INDEX IF NOT EXISTS idx_assets_status ON public.assets(status);
CREATE INDEX IF NOT EXISTS idx_assets_date ON public.assets(date DESC);

-- 3. Row Level Security (RLS) 설정
-- 웹앱에서 익명(anon) 키로 직접 읽기/쓰기가 가능하도록 정책을 활성화합니다.
ALTER TABLE public.assets ENABLE ROW LEVEL SECURITY;

-- 기존 정책이 있다면 삭제 후 재생성
DROP POLICY IF EXISTS "Enable read access for all users" ON public.assets;
DROP POLICY IF EXISTS "Enable insert access for all users" ON public.assets;
DROP POLICY IF EXISTS "Enable update access for all users" ON public.assets;
DROP POLICY IF EXISTS "Enable delete access for all users" ON public.assets;

-- 읽기(SELECT) 허용 정책
CREATE POLICY "Enable read access for all users"
ON public.assets FOR SELECT
USING (true);

-- 등록(INSERT) 허용 정책
CREATE POLICY "Enable insert access for all users"
ON public.assets FOR INSERT
WITH CHECK (true);

-- 수정(UPDATE) 허용 정책
CREATE POLICY "Enable update access for all users"
ON public.assets FOR UPDATE
USING (true)
WITH CHECK (true);

-- 삭제(DELETE) 허용 정책
CREATE POLICY "Enable delete access for all users"
ON public.assets FOR DELETE
USING (true);

-- 4. 수정 시각(updated_at) 자동 갱신 트리거 설정
CREATE OR REPLACE FUNCTION public.handle_updated_at()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = NOW();
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trigger_assets_updated_at ON public.assets;
CREATE TRIGGER trigger_assets_updated_at
BEFORE UPDATE ON public.assets
FOR EACH ROW
EXECUTE FUNCTION public.handle_updated_at();

-- 5. 기본 실무 샘플 데이터 초기 입력 (선택 사항)
INSERT INTO public.assets (code, name, category, date, quantity, price, department, "user", location, status, serial, remarks)
VALUES
    ('AST-IT-2024-001', 'MacBook Pro 16인치 M3 Pro', 'IT기기', '2024-01-15', 1, 3490000, '개발팀', '김민수 수석', '본사 5층 개발실 A-12', '사용중', 'C02G99A0MD6R', '애플케어 플러스 적용 (~2027)'),
    ('AST-IT-2024-002', 'Dell UltraSharp 32인치 4K', 'IT기기', '2024-01-20', 2, 950000, '디자인팀', '이서연 팀장', '본사 4층 디자인룸', '사용중', 'CN-0R289D-74445', '듀얼 모니터 구성'),
    ('AST-FUR-2023-015', '허먼밀러 에어론 체어 B사이즈', '사무가구', '2023-08-10', 5, 1850000, '경영지원팀', '임원실/총무팀', '본사 6층 임원실', '사용중', 'HM-AER-202308', '12년 본사 워런티 보증'),
    ('AST-IT-2023-088', 'LG 그램 15인치 업무용 노트북', 'IT기기', '2023-05-12', 1, 1650000, '영업팀', '박준혁 과장', '영업본부 외근용', '수리중', 'LG15Z90R-GA56K', '액정 패널 AS센터 접수 중'),
    ('AST-ELC-2023-004', '삼성 비스포크 공기청정기', '전자제품', '2023-03-02', 3, 680000, '공용/총무', '사옥 공용', '대회의실 및 라운지', '사용중', 'AX53A9310GGD', '필터 정기교체(매년 4월)'),
    ('AST-FUR-2022-042', '데스커 전동 모션데스크 1600', '사무가구', '2022-11-05', 1, 520000, '개발팀', '미배정', '지하 1층 창고', '유휴(보관)', 'DSK-MD-1600W', '신규 입사자 배정 대기'),
    ('AST-CAR-2022-001', '현대 아이오닉5 (법인차량)', '차량운반구', '2022-06-18', 1, 54000000, '경영지원팀', '법인 공용', '지하 2층 주차장', '사용중', '123가 4567', '업무운행일지 작성 필수'),
    ('AST-IT-2021-030', 'HP 고속 복합기 MFP-E82660', 'IT기기', '2021-09-01', 1, 3200000, '공용/총무', '전사 공용', '본사 4층 복사실', '사용중', 'HPE82660-KR003', '월 렌탈 유지보수 계약'),
    ('AST-IT-2020-011', '구형 데스크톱 본체 (i5-8세대)', 'IT기기', '2020-04-10', 2, 850000, '개발팀', '불용자산', '지하 창고 폐기구역', '폐기예정', 'SYS-OLD-2020', '디가우징 후 폐기 예정')
ON CONFLICT (code) DO NOTHING;
