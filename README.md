# 판매 대시보드

결산 데이터(월별 엑셀 / 일별 CSV)를 업로드하면 자동으로 파싱·시각화하는 사내용 웹 대시보드입니다. Streamlit Cloud + Supabase 조합으로 운영됩니다.

---

## 📦 프로젝트 구조

```
aram_daily_dashboard/
├── app.py                   # 메인 앱 (Streamlit)
├── requirements.txt         # 파이썬 패키지 목록
├── README.md                # 이 문서
├── .gitignore               # GitHub에 안 올릴 파일 목록
│
├── .streamlit/
│   └── secrets.toml         # 로컬 전용 (GitHub에 절대 올리지 않음)
│
└── logo/                    # 업체 로고 이미지 - supabase로 이관 (보안유지) 
```

### 클라우드 인프라

```
[Streamlit Cloud] ──── 호스팅 (퍼블릭 repo + 앱 비밀번호 보호)
        │
        ├── 파일 업로드 ──▶ [Supabase Storage]   uploads 버킷(원본 xls/xlsx/csv), logos 버킷(업체 로고)
        │                          │
        │                     앱이 읽어 파싱
        │                          │
        └── 입력/조회 ──────▶ [Supabase Postgres]
                                   ├── monthly   월별 판매금액
                                   ├── daily     일별 누적 매출
                                   ├── targets   목표 (LG생활건강 · 월별)
                                   ├── memos     미납현황
                                   ├── notices   공지사항
                                   └── uploaded_files  처리 완료 파일 기록 (중복 업로드 방지)
```

---

## 🖥 주요 기능 (메뉴 4개)

> 화면 상단 가로 라디오 버튼(`st.radio`)으로 이동합니다. (`st.tabs`는 렌더링 오류가 있어 교체)

### 📊 대시보드
- **업체별 카드 (3×3 그리드)** — 일별 누적 매출액 큰 순으로 자동 정렬, 순위 표시
- **카드 내용** — 일별 누적 매출액(상단) + 월별 막대그래프(하단, 각 막대 위에 금액 표시)
  - 막대그래프 위 **◀ 연도 ▶** 버튼으로 연도 이동 (기본값: 올해, 올해 이후로는 이동 불가)
  - 1~6월 / 7~12월 두 줄로 나눠 표시
- **LG생활건강 카드**에는 목표 달성률(%)이 일별 매출 옆에 같이 표시됨
- **도넛 차트** — 전체 비중. 업체명·비중%가 도넛 바깥에 라벨로 표시, 하단 범례에도 비중%
- **공지사항** — 날짜별 공지 작성·저장 (`notices` 테이블)
- **미납현황 입력** — 도넛 아래에서 날짜 선택 + 업체 체크 + 조치항목 입력 후 **💾 미납현황 저장** (삼립은 제외)
  - `st.form`으로 묶여 있어 입력 중에는 화면이 새로고침되지 않고, 저장 버튼을 눌러야 반영됨

### 📁 파일 업로드
- 업체 선택 → 종류(월별/일별) 선택 → 파일 드래그앤드롭 (여러 개 가능)
- 자동 처리: Storage에 원본 저장 + 파싱해서 DB에 upsert
- 같은 업체·같은 달(또는 같은 날짜) 파일을 다시 올리면 덮어쓰기
- **같은 파일명**은 "이미 처리된 파일"로 건너뜀 → 다시 처리하려면 `uploaded_files` 테이블에서 해당 행 삭제 후 재업로드
  - 파싱에 실패한 파일은 기록되지 않으므로 고친 뒤 같은 이름으로 다시 올려도 됨
- **남양유업**은 파일 2개(`060`, `b22`)를 같이 올리면 같은 달/날짜끼리 자동 합산

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
- **확장자가 아니라 파일 내용(첫 바이트)으로 형식을 판별**해서 읽음

| 파일 앞부분 | 실제 형식 | 읽는 방법 |
|---|---|---|
| `PK` | 진짜 xlsx | `openpyxl` |
| `D0 CF 11 E0` | 진짜 xls | `xlrd` → 실패하면 `calamine` 으로 재시도 |
| 그 외 | 웹페이지(HTML) 형식 | `pd.read_html` (cp949) |

- 복지단 파일은 이름만 `.xlsx`이고 내용은 xls/HTML인 경우가 있음 → 이 방식으로 자동 처리
- 비표준 xls는 xlrd가 못 읽는 경우가 있어 `calamine` 으로 한 번 더 시도
- 파일에서 "YYYY년 MM월" 찾아 월 추출
- **"판매금액" 글자가 있는 열을 자동으로 찾아** 그 열을 사용 (열 번호 고정 아님)
- "금액 계" 행의 다음 줄에 있는 판매금액 사용
- 금액은 `_num()`으로 쉼표를 제거한 뒤 정수 변환 (HTML 형식은 `"12,345"` 처럼 글자로 읽히기 때문)
- **삼립**은 예외: 지정된 16개 제품코드만 합산
  - `220176, 220177, 220178, 230226, 230229, 230232, 240218, 240219, 240220, 250225, 250226, 250228, 260205, 260206, 260207, 260209`

### 일별 CSV (CP949 인코딩)
- 파일명 앞 `YYYYMMDD` = 조회일 (그 달 1일~조회 전일까지의 **누적**)
- 한 행 = 한 품목, **매출액 = 판매수량 × 공급단가**
- **기준일(YYYYMMDD) 열을 앵커로** 위치 계산: 판매수량 = 기준일 − 3, 공급단가 = 기준일 − 1
  - 업체마다 열 개수가 다름(LG생활건강 19~21열, 제품명·지역명에 쉼표) → 뒤에서 세는 방식은 금액이 크게 틀어져서 폐기
  - 앵커를 못 찾거나 단가가 음수인 행은 제외하고 화면에 "제외 N행" 표시
- **업체 검증**: 파일 3번째 열의 업체코드(소문자 비교)가 `VENDOR_TOKENS`와 맞지 않으면 저장하지 않음
  - 파일명에 `YYYYMMDD_코드_` 형태가 있으면 파일 내용 코드와도 대조
  - `VENDOR_TOKENS` 값은 **소문자**로 적어야 함 (예: `9hh`, `9bs`)
- 합계가 음수면 저장하지 않음 (파일 형식 이상 신호)
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

앱이 추가로 사용하는 테이블 (앱 코드에서 쓰는 컬럼 기준):

| 테이블 | 컬럼 | 비고 |
|---|---|---|
| `notices` | 날짜, 공지내용 | 날짜 기준 upsert (날짜 UNIQUE 필요) |
| `uploaded_files` | 업체, 종류(`monthly`/`daily`), 파일명, 월 또는 날짜 | 같은 업체·종류·파일명이면 업로드 건너뜀 |

> 모든 테이블은 RLS를 꺼야 합니다: `ALTER TABLE <테이블> DISABLE ROW LEVEL SECURITY;`

### Storage 경로 (`uploads` 버킷)

```
uploads/
├── lg/monthly/<파일명>.xls(x)
├── lg/daily/<파일명>.csv
├── samlip/monthly/...
└── ... (영문 코드 폴더만 허용 — 한글 경로 불가)
```

업체명 → 영문 코드 매핑:
- LG생활건강 → `lg`
- 비알코리아 → `br`
- 에너자이저 → `energizer`
- 라벨리 → `lavelee`
- 메디카 → `medica`
- 남양유업 → `namyang`
- 나사라 → `nasara`
- 삼립 → `samlip`
- 티젠 → `teazen`

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
python-calamine
lxml
supabase
streamlit-autorefresh
```

- `python-calamine`: xlrd가 못 읽는 비표준 xls 대비
- `lxml`: 웹페이지(HTML) 형식 엑셀 읽기용
- ⚠️ Streamlit Cloud는 `requirements.txt`가 바뀌면 환경을 새로 만들면서 파이썬 버전이 올라갈 수 있음. 배포 후 패키지 에러가 나면 버전을 명시(`패키지==버전`)해서 고정

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
| 월별 업로드 `unpack requires a buffer of 2 bytes` | xlrd가 못 읽는 비표준 xls → 코드에서 `calamine`으로 자동 재시도. 그래도 안 되면 엑셀에서 열어 **다른 이름으로 저장 → xlsx** |
| 월별 업로드 `File is not a zip file` | 이름만 `.xlsx`이고 실제는 xls/HTML → 코드가 첫 바이트로 형식을 판별해 자동 처리. 그래도 안 되면 엑셀에서 **다른 이름으로 저장 → xlsx** (확장자만 바꾸면 안 됨) |
| 월별 업로드 "월/판매금액을 못 찾았습니다" | 파일에 "YYYY년 MM월" 또는 "판매금액"/"금액 계" 글자가 없음 → 원본 양식 확인 |
| 일별 업로드 "'OOO' 선택했는데 파일은 …" | 업체 선택 실수 또는 `VENDOR_TOKENS` 코드가 대문자로 적힘 → 소문자로 수정 |
| "이미 처리된 파일입니다" | 같은 파일명이 `uploaded_files`에 있음 → 다시 처리하려면 해당 행 삭제 후 재업로드 |

---

## 📜 변경 이력

- **v1**: 로컬 Streamlit + JSON 파일 (사내 PC에서 실행)
- **v2**: Streamlit Cloud + Supabase 전환, 탭 구조, 미납현황 RDB
- **v3**: 삼립 제품코드 필터링, 한국 시간 표시, 콤마 포맷
- **v4**: 카드 정렬(일별 매출 순), 도넛 바깥 라벨, 목표 설정 별도 탭
- **v5**: 일별 CSV 기준일 앵커 파싱, 업체코드 검증, 남양유업 2파일 합산, 공지사항, 중복 업로드 방지(`uploaded_files`), 메뉴를 `st.radio`로 교체
- **v6** (2026-09): 월별 막대그래프 연도 이동(◀ ▶), 미납현황 입력을 `st.form`으로 변경, 로고를 Supabase `logos` 버킷으로 이관
- **v7** (2026-11): 월별 엑셀 형식 자동 판별(xlsx/xls/HTML), 비표준 xls `calamine` 재시도, 금액 쉼표 처리
