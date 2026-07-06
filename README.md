# 판매 대시보드

복지단 결산 데이터(월별 엑셀 / 일별 CSV)를 업로드하면 자동으로 파싱·시각화하는 사내용 웹 대시보드입니다. Streamlit Cloud + Supabase 조합으로 운영됩니다.

---

## 📦 프로젝트 구조

```
sales-dashboard/
├── app.py                   # 메인 앱 (Streamlit)
├── requirements.txt         # 파이썬 패키지 목록
├── README.md                # 이 문서
├── .gitignore               # GitHub에 안 올릴 파일 목록
│
├── .streamlit/
│   └── secrets.toml         # 로컬 전용 (GitHub에 절대 올리지 않음)
│
└── logo/                    # 업체 로고 이미지
    ├── LG생활건강.png
    ├── 비알코리아.png
    ├── 에너자이저.png
    ├── 라벨리.png
    ├── 메디카.png
    ├── 남양유업.png
    ├── 나사라.png
    ├── 삼립.png
    └── 티젠.png
```

### 클라우드 인프라

```
[Streamlit Cloud] ──── 호스팅 (퍼블릭 repo + 앱 비밀번호 보호)
        │
        ├── 파일 업로드 ──▶ [Supabase Storage]   uploads 버킷, 원본 xls/csv
        │                          │
        │                     앱이 읽어 파싱
        │                          │
        └── 입력/조회 ──────▶ [Supabase Postgres]
                                   ├── monthly   월별 판매금액
                                   ├── daily     일별 누적 매출
                                   ├── targets   목표 (LG생활건강 · 월별)
                                   └── memos     미납현황
```

---

## 🖥 주요 기능 (탭 4개)

### 📊 대시보드
- **업체별 카드 (3×3 그리드)** — 일별 누적 매출액 큰 순으로 자동 정렬, 순위 표시
- **카드 내용** — 일별 누적 매출액(상단) + 월별 막대그래프(하단, 각 막대 위에 금액 표시)
- **LG생활건강 카드**에는 목표 달성률(%)이 일별 매출 옆에 같이 표시됨
- **도넛 차트** — 전체 비중. 업체명·비중%가 도넛 바깥에 라벨로 표시, 하단 범례에도 비중%
- **미납현황 입력** — 도넛 아래에서 날짜 선택 + 9개 업체 체크 + 조치항목, 체크된 행만 DB에 저장

### 📁 파일 업로드
- 업체 선택 → 종류(월별/일별) 선택 → 파일 드래그앤드롭 (여러 개 가능)
- 자동 처리: Storage에 원본 저장 + 파싱해서 DB에 upsert
- 같은 업체·같은 달(또는 같은 날짜) 파일을 다시 올리면 덮어쓰기

### 💰 매출 이력
- 월별 매출 / 일별 누적 매출 / 미납현황 중 선택
- 업체·기간 필터 → 표로 조회, 합계 표시, CSV 다운로드
- 시각(업로드시각·수정시각)은 한국 시간(KST)으로 자동 변환

### 🎯 목표 설정
- LG생활건강 월별 목표 금액 입력·저장
- 저장된 목표 이력 표로 조회 (옛 달 목표 자동 보존)

---

## 🛠 데이터 처리 규칙

### 월별 엑셀 (.xls / .xlsx)
- 파일에서 "YYYY년 MM월" 찾아 월 추출
- "금액 계" 행의 다음 줄에 있는 **판매금액 (col 25)** 사용
- **삼립**은 예외: 지정된 16개 제품코드만 합산
  - `220176, 220177, 220178, 230226, 230229, 230232, 240218, 240219, 240220, 250225, 250226, 250228, 260205, 260206, 260207, 260209`

### 일별 CSV (CP949 인코딩)
- 파일명 앞 `YYYYMMDD` = 조회일 (그 달 1일~조회 전일까지의 **누적**)
- 한 행 = 한 품목, **매출액 = 판매수량 × 공급단가**
- 컬럼 위치: 판매수량 = 뒤에서 7번째, 공급단가 = 뒤에서 5번째 (제품명에 쉼표가 들어가도 안전)
- 삼립은 위 16개 제품코드만 합산

### Supabase 테이블 구조

```sql
CREATE TABLE monthly (
    id BIGSERIAL PRIMARY KEY,
    업체 TEXT NOT NULL,
    월 TEXT NOT NULL,                -- "2026-04"
    판매금액 BIGINT NOT NULL,
    업로드시각 TIMESTAMPTZ DEFAULT NOW(),
    UNIQUE(업체, 월)
);

CREATE TABLE daily (
    id BIGSERIAL PRIMARY KEY,
    업체 TEXT NOT NULL,
    날짜 TEXT NOT NULL,               -- "2026-06-17"
    매출액 BIGINT NOT NULL,
    업로드시각 TIMESTAMPTZ DEFAULT NOW(),
    UNIQUE(업체, 날짜)
);

CREATE TABLE targets (
    업체 TEXT NOT NULL,
    월 TEXT NOT NULL,
    목표금액 BIGINT NOT NULL,
    PRIMARY KEY(업체, 월)
);

CREATE TABLE memos (
    id BIGSERIAL PRIMARY KEY,
    날짜 TEXT NOT NULL,
    업체 TEXT NOT NULL,
    미납여부 BOOLEAN DEFAULT FALSE,
    미납내용 TEXT DEFAULT '',
    조치항목 TEXT DEFAULT '',
    수정시각 TIMESTAMPTZ DEFAULT NOW(),
    UNIQUE(날짜, 업체)
);
```

### Storage 경로 (`uploads` 버킷)

```
uploads/
├── lg/monthly/<파일명>.xls
├── lg/daily/<파일명>.csv
├── samlip/monthly/...
└── ... (영문 코드 폴더만 허용 — 한글 경로 불가)
```

업체명 → 영문 코드 매핑:
- LG생활건강 → `lg`
- 비알코리아 → `br`
- 에너자이저 → `energizer`
- 라벨리 → `labelly`
- 메디카 → `medica`
- 남양유업 → `namyang`
- 나사라 → `nasara`
- 삼립 → `samlip`
- 티젠 → `tjeen`

---

## ⚙️ 로컬 실행

### 1) 환경 준비

```bash
pip install -r requirements.txt
```

`requirements.txt` 내용:
```
streamlit
pandas
altair
xlrd
openpyxl
supabase
```

### 2) Supabase 키 등록

`app.py` 와 같은 폴더에 `.streamlit/secrets.toml` 파일을 만들고 다음 내용을 넣습니다.

```toml
SUPABASE_URL = "https://여러분프로젝트.supabase.co"
SUPABASE_KEY = "eyJhbGc..."
APP_PASSWORD = "사내공유비밀번호"
```

> ⚠️ 이 파일은 절대 GitHub에 올리지 않습니다 (`.gitignore` 에 등록됨).

### 3) 실행

```bash
streamlit run app.py
```

브라우저가 열리고 비밀번호 입력 화면이 나옵니다.

---

## 🚀 클라우드 배포 (Streamlit Cloud)

1. GitHub **퍼블릭 repo**에 `app.py`, `requirements.txt`, `.gitignore`, `logo/` 폴더 업로드
   - `.streamlit/secrets.toml` 은 절대 올리지 않음 (`.gitignore` 가 막아줌)
2. https://share.streamlit.io 에서 GitHub 연동 → repo 선택 → Main file `app.py`
3. **Advanced settings → Secrets** 에 위 `secrets.toml` 내용을 그대로 붙여넣음
4. **Deploy!** 클릭. 2~5분 후 배포 완료
5. 접속 URL과 비밀번호를 직원들에게 공유

---

## 📊 운영 가이드

### 사용량 확인
- **Supabase**: 대시보드 → Project Settings → Usage
  - 무료 한도: DB 500MB, Storage 1GB, Egress 5GB/월
- **Streamlit Cloud**: 앱 페이지 → Analytics
  - 7일간 미접속 시 슬립 (평일 매일 보는 패턴이면 영향 없음)

### 데이터 삭제 (Supabase SQL Editor)

```sql
-- 일별 매출 전체 삭제
DELETE FROM daily;

-- 미납현황 전체 삭제
DELETE FROM memos;

-- 특정 업체·기간만 삭제
DELETE FROM daily WHERE 업체 = 'LG생활건강' AND 날짜 < '2026-06-01';

-- 업체명 통일 (예: 옛 이름 → 새 이름)
UPDATE monthly SET 업체 = 'LG생활건강' WHERE 업체 = '엘지';
UPDATE daily   SET 업체 = 'LG생활건강' WHERE 업체 = '엘지';
UPDATE targets SET 업체 = 'LG생활건강' WHERE 업체 = '엘지';
UPDATE memos   SET 업체 = 'LG생활건강' WHERE 업체 = '엘지';
```

> 삭제는 되돌릴 수 없습니다. 큰 작업 전에는 Table Editor에서 **Export → CSV** 로 백업하세요.

### 데이터가 화면에 안 보일 때
- 화면 상단 **🔄 새로고침** 버튼 클릭 (캐시 30초)
- 그래도 안 보이면 Supabase Table Editor에서 행이 실제로 들어갔는지 확인

---

## 🔒 보안 메모

- **퍼블릭 repo + 앱 비밀번호** 방식으로 운영 중
- 코드 자체엔 민감 정보 없음 (Supabase 키는 Streamlit Secrets에만 존재)
- 매출 데이터는 Supabase Postgres에 저장되고, anon key + 앱 비밀번호로 보호됨
- 비밀번호는 3~6개월에 한 번 변경 권장

### RLS (Row Level Security) 정책
- 테스트 단계에선 RLS를 끄고 사용 중
- 더 엄밀한 보안이 필요하면 Supabase 정책으로 IP 제한 또는 service_role 키 + 서버 사이드 인증 검토

---

## 🆘 자주 막히는 지점

| 증상 | 원인 / 해결 |
|---|---|
| `Invalid key: <한글경로>` | Storage 경로에 한글이 들어감 → COMPANY_CODES 매핑 확인 |
| `row violates row-level security policy` | RLS가 켜져 있음 → SQL Editor에서 `ALTER TABLE ... DISABLE ROW LEVEL SECURITY;` |
| `KeyError: '엘지'` 등 | DB의 옛 업체명이 새 이름과 다름 → UPDATE 쿼리로 통일 |
| 도넛이 깨져 보임 | 음수 합계 업체가 섞임 → `overall[overall["판매금액"] > 0]` 필터 (이미 적용됨) |
| 배포 후 `ModuleNotFoundError` | `requirements.txt` 에 빠진 패키지 있음 |
| `secrets.toml` 관련 에러 | Streamlit Cloud의 Secrets 등록 누락 → Manage app → Settings → Secrets |

---

## 📜 변경 이력

- **v1**: 로컬 Streamlit + JSON 파일 (사내 PC에서 실행)
- **v2**: Streamlit Cloud + Supabase 전환, 탭 구조, 미납현황 RDB
- **v3**: 삼립 제품코드 필터링, 한국 시간 표시, 콤마 포맷
- **v4**: 카드 정렬(일별 매출 순), 도넛 바깥 라벨, 목표 설정 별도 탭
