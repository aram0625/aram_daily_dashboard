# -*- coding: utf-8 -*-
"""
판매 대시보드 - Supabase 연동 버전 (Streamlit Cloud 배포용)

탭 구성:
  📊 대시보드   : 업체별 카드(월별/일별) + 도넛(전체 비중) + 공지 + 미납현황
  📁 파일 업로드: 월별 엑셀 / 일별 CSV 다중 업로드 → Supabase에 저장
  💰 매출 이력 : monthly/daily/memos/notices를 표로 조회
  🎯 목표 설정 : LG생활건강 월별 목표 (이력 자동 보존)
"""

import io
import re
import csv
import os
import sys
import unicodedata
import datetime

import pandas as pd
import altair as alt
import streamlit as st
from supabase import create_client
from zoneinfo import ZoneInfo

# 날짜 
today = datetime.datetime.now(ZoneInfo("Asia/Seoul")).strftime("%Y년 %m월 %d일")

st.set_page_config(page_title="판매 대시보드", page_icon="📊", layout="wide")

# 글씨 크기 줄이기
st.markdown("""
<style>
[data-testid="stMetricValue"] { font-size: 22px !important; }
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------
# 접근 비밀번호
# ----------------------------------------------------------------------
def check_password():
    if st.session_state.get("auth_ok"):
        return True

    st.title("🔒 아람비즈 가는길") 
    st.caption("접근 비밀번호를 입력하세요.")
    pw = st.text_input("비밀번호", type="password",
                       label_visibility="collapsed")
    col_btn, _ = st.columns([1, 4])
    if col_btn.button("로그인", use_container_width=True):
        if pw == st.secrets.get("APP_PASSWORD", ""):
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.error("비밀번호가 틀렸습니다.")
    return False


if not check_password():
    st.stop()


# ----------------------------------------------------------------------
# Supabase 연결
# ----------------------------------------------------------------------
@st.cache_resource
def get_supabase():
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])

sb = get_supabase()
BUCKET = "uploads"


# ----------------------------------------------------------------------
# 고정 설정
# ----------------------------------------------------------------------
COMPANIES = ["LG생활건강", "비알코리아", "에너자이저", "라벨리", "메디카",
             "남양유업", "나사라", "삼립", "티젠"]

PALETTE = ["#a44646", "#93f0e3", "#5e5e5e", "#f9b3b3", "#77f2ae",
           "#e26363", "#e6b189", "#cbc4f7", "#f9e3a3"]
COMPANY_COLORS = {n: PALETTE[i % len(PALETTE)] for i, n in enumerate(COMPANIES)}

COMPANY_CODES = {
    "LG생활건강": "lg", "비알코리아": "br", "에너자이저": "energizer",
    "라벨리": "lavelee", "메디카": "medica", "남양유업": "namyang",
    "나사라": "nasara", "삼립": "samlip", "티젠": "teazen",
}

SELECTED_CODES = {
    "삼립": {"220176","220177","220178","230226","230229","230232","240218",
             "240219","240220","250225","250226","250228","260205","260206",
             "260207","260209"},
}


def app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE = app_dir()
LOGO_DIR = os.path.join(BASE, "logo")

# 로고 
@st.cache_data(ttl=3600)   # 1시간 캐시 (이미지는 잘 안 바뀜)
def logo_for(company: str):
    """Supabase Storage(logos 버킷)에서 로고 URL을 가져옴."""
    code = COMPANY_CODES.get(company)
    if not code:
        return None
    try:
        # Public 버킷이라 URL만 반환 (Streamlit이 알아서 다운로드)
        url = sb.storage.from_("logos").get_public_url(f"{code}.png")
        return url
    except Exception:
        return None


def safe_path(s: str) -> str:
    s = unicodedata.normalize("NFKC", s)
    s = re.sub(r"[^A-Za-z0-9._-]", "_", s)
    return s.strip("_") or "x"


def company_code(name: str) -> str:
    return COMPANY_CODES.get(name, safe_path(name))


# ----------------------------------------------------------------------
# 파서
# ----------------------------------------------------------------------
def parse_monthly(buf, filename, company=None):
    ext = filename.lower().split(".")[-1]
    engine = "xlrd" if ext == "xls" else "openpyxl"
    df = pd.read_excel(buf, engine=engine, header=None)
    nrows = len(df)

    month = None
    for i in range(nrows):
        for j in range(df.shape[1]):
            v = df.iloc[i, j]
            if isinstance(v, str):
                m = re.search(r"(\d{4})\s*년\s*(\d{1,2})\s*월", v)
                if m:
                    month = f"{int(m.group(1))}-{int(m.group(2)):02d}"
                    break
        if month:
            break

    sales_col = None
    for i in range(nrows):
        for j in range(df.shape[1]):
            if df.iloc[i, j] == "판매금액":
                sales_col = j
                break
        if sales_col is not None:
            break

    total = None
    selected = SELECTED_CODES.get(company)
    if selected and sales_col is not None:
        s = 0
        for i in range(nrows):
            code = df.iloc[i, 0]
            if pd.isna(code):
                continue
            if str(code).strip() in selected and i + 1 < nrows:
                v = df.iloc[i + 1, sales_col]
                if pd.notna(v):
                    s += int(v)
        total = s
    else:
        for i in range(nrows):
            v = df.iloc[i, 0]
            if isinstance(v, str) and v.replace(" ", "") == "금액계":
                if sales_col is not None and i + 1 < nrows:
                    val = df.iloc[i + 1, sales_col]
                    if pd.notna(val):
                        total = int(val)
                break

    return month, total


def parse_daily(buf, filename, company=None):
    m = re.search(r"(20\d{6})", filename)
    if not m:
        return None, None
    s = m.group(1)
    date = f"{s[:4]}-{s[4:6]}-{s[6:8]}"

    selected = SELECTED_CODES.get(company)
    text = buf.read().decode("cp949", errors="replace")
    total = 0.0
    for row in csv.reader(io.StringIO(text)):
        if len(row) < 7:
            continue
        if selected:
            code = row[4].strip() if len(row) > 4 else ""
            if code not in selected:
                continue
        try:
            qty = float(row[-7])
            price = float(row[-5])
        except ValueError:
            continue
        total += qty * price
    return date, int(round(total))


# ----------------------------------------------------------------------
# DB 조회 헬퍼
# ----------------------------------------------------------------------
@st.cache_data(ttl=30)
def fetch_monthly():
    rows = sb.table("monthly").select("*").order("월").execute().data
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=["업체", "월", "판매금액"])


@st.cache_data(ttl=30)
def fetch_daily():
    rows = sb.table("daily").select("*").order("날짜").execute().data
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=["업체", "날짜", "매출액"])


@st.cache_data(ttl=30)
def fetch_targets():
    rows = sb.table("targets").select("*").execute().data
    out = {}
    for r in rows or []:
        out.setdefault(r["업체"], {})[r["월"]] = int(r["목표금액"])
    return out


@st.cache_data(ttl=30)
def fetch_memos(date_key=None):
    q = sb.table("memos").select("*")
    if date_key:
        q = q.eq("날짜", date_key)
    rows = q.order("날짜", desc=True).order("업체").execute().data
    return pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=["날짜", "업체", "미납여부", "미납내용", "조치항목"]
    )


@st.cache_data(ttl=30)
def fetch_notices(date_key=None):
    q = sb.table("notices").select("*")
    if date_key:
        q = q.eq("날짜", date_key)
    rows = q.order("날짜", desc=True).execute().data
    return pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=["날짜", "공지내용", "수정시각"]
    )


def invalidate_cache():
    fetch_monthly.clear()
    fetch_daily.clear()
    fetch_targets.clear()
    fetch_memos.clear()
    fetch_notices.clear()


# ======================================================================
# 헤더
# ======================================================================
title_col, btn_col = st.columns([7, 1], vertical_alignment="bottom")
with title_col:
    st.markdown(
        f"<div style='display:flex; align-items:baseline; gap:15px'>"
        f"<span style='font-size:2.5rem; font-weight:700'>아람비즈 가는길</span>"
        f"<span style='font-size:24px; font-weight:600; color:#444'>{today}</span>"
        f"</div>",
        unsafe_allow_html=True,
    )
with btn_col:
    if st.button("🔄 새로고침"):
        invalidate_cache()
        st.rerun()
        
monthly_df = fetch_monthly()
daily_df = fetch_daily()
TARGETS = fetch_targets()

if not monthly_df.empty:
    overall = (monthly_df.groupby("업체", as_index=False)["판매금액"].sum()
                         .sort_values("판매금액", ascending=False))
else:
    overall = pd.DataFrame(columns=["업체", "판매금액"])


# ======================================================================
# 탭
# ======================================================================
tab_dash, tab_upload, tab_hist, tab_target = st.tabs(
    ["📊 대시보드", "📁 파일 업로드", "💰 매출 이력", "🎯 목표 설정"]
)

# ----------------------------------------------------------------------
# 📊 대시보드
# ----------------------------------------------------------------------
with tab_dash:
    main_left, main_right = st.columns([3, 2])

    # ----- 왼쪽: 카드 그리드 -----
    with main_left:
        st.subheader("업체별 현황")

        if not daily_df.empty:
            latest_daily = (daily_df.sort_values("날짜")
                                    .groupby("업체").tail(1)
                                    .set_index("업체")["매출액"].to_dict())
        else:
            latest_daily = {}

        ordered = sorted(COMPANIES, key=lambda c: latest_daily.get(c, 0),
                         reverse=True)
        rank_of = {n: i + 1 for i, n in enumerate(ordered)}

        cols_per_row = 3
        for i in range(0, len(ordered), cols_per_row):
            row = st.columns(cols_per_row)
            for col, name in zip(row, ordered[i:i + cols_per_row]):
                with col:
                    with st.container(border=True):
                        lp = logo_for(name)
                        head = st.columns([1, 3], vertical_alignment="center")
                        with head[0]:
                            if lp:
                                st.image(lp, width=40)
                            else:
                                st.markdown(
                                    "<div style='font-size:30px'>🟥</div>",
                                    unsafe_allow_html=True,
                                )
                        head[1].markdown(f"**{rank_of[name]}위 · {name}**")

                        sdf = monthly_df[monthly_df["업체"] == name].sort_values("월")
                        color = COMPANY_COLORS.get(name, "#A50034")

                        # 일별 누적 매출액 + 달성률
                        ddf = (daily_df[daily_df["업체"] == name]
                               .sort_values("날짜").reset_index(drop=True))
                        if ddf.empty:
                            st.metric("일별 누적 매출액 (원)", "데이터 없음")
                        else:
                            latest_date = ddf["날짜"].iloc[-1]
                            latest_total = int(ddf["매출액"].iloc[-1])
                            month_key = latest_date[:7]
                            target = TARGETS.get(name, {}).get(month_key)

                            st.markdown(
                                f"<span style='font-size:13px; color:gray'>"
                                f"{latest_date}</span>",
                                unsafe_allow_html=True,
                            )

                            m_col, t_col = st.columns([2, 1], vertical_alignment="center")
                            with m_col:
                                st.metric(
                                    "일별 누적 매출액 (원)",
                                    f"{latest_total:,}",
                                    help=(f"{month_key} 목표: {target:,} 원"
                                          if target else
                                          f"{latest_date} 기준 · 그 달 1일부터 "
                                          f"조회 전일까지 누적"),
                                )
                            with t_col:
                                if target:
                                    pct = latest_total / target * 100
                                    st.markdown(
                                        f"<div style='font-size:20px; "
                                        f"font-weight:bold; color:#1f77b4; "
                                        f"text-align:center; padding-top:20px'>"
                                        f"{pct:.1f}%<br>"
                                        f"<span style='font-size:12px; "
                                        f"color:gray; font-weight:normal'>달성</span>"
                                        f"</div>",
                                        unsafe_allow_html=True,
                                    )

                        # 월별 막대그래프 (높이 키움: 170 → 240)
                        if sdf.empty:
                            st.caption("월별 데이터 없음")
                        else:
                            st.caption("월별 판매금액 (원)")
                            sdf2 = sdf.assign(월라벨=sdf["월"].str[-2:])
                            base_m = alt.Chart(sdf2).encode(
                                x=alt.X("월라벨:N", sort=None, title=None,
                                        axis=alt.Axis(labelAngle=0,
                                                      labelFontSize=18)),
                            )
                            bar = base_m.mark_bar(
                                color=color, cornerRadiusTopLeft=3,
                                cornerRadiusTopRight=3, size=35,
                            ).encode(
                                y=alt.Y("판매금액:Q", title=None, axis=None),
                                tooltip=[alt.Tooltip("월:N"),
                                         alt.Tooltip("판매금액:Q", format=",")],
                            )
                            text = base_m.mark_text(
                                dy=-8, fontSize=11, color="#333",
                            ).encode(
                                y=alt.Y("판매금액:Q"),
                                text=alt.Text("판매금액:Q", format=","),
                            )
                            st.altair_chart(
                                alt.layer(bar, text).properties(height=280),
                                use_container_width=True,
                            )
             # 행 사이 여백
            st.write("")
    # ----- 오른쪽: 도넛 + 공지 + 미납 -----
    with main_right:
        st.subheader("전체 비중")

        overall_d = overall[overall["판매금액"] > 0].copy()

        if overall_d.empty:
            st.info("표시할 업체 데이터가 없습니다.")
        else:
            total_all = overall_d["판매금액"].sum()
            overall_d["비중"] = overall_d["판매금액"] / total_all * 100

            overall_d["라벨"] = overall_d.apply(
                lambda r: f"{r['업체']} {r['비중']:.1f}%" if r["비중"] >= 2 else "",
                axis=1,
            )

            overall_d["원래업체"] = overall_d["업체"]
            overall_d["업체"] = overall_d.apply(
                lambda r: f"{r['원래업체']} ({r['비중']:.1f}%)", axis=1,
            )

            base = alt.Chart(overall_d).encode(
                theta=alt.Theta("판매금액:Q", stack=True),
                tooltip=[
                    alt.Tooltip("업체:N"),
                    alt.Tooltip("판매금액:Q", format=","),
                    alt.Tooltip("비중:Q", format=".1f", title="비중(%)"),
                ],
            )

            arc = base.mark_arc(
                innerRadius=65, outerRadius=110,
                stroke="white", strokeWidth=1,
            ).encode(
                color=alt.Color(
                    "업체:N",
                    scale=alt.Scale(
                        domain=overall_d["업체"].tolist(),
                        range=[COMPANY_COLORS[c]
                               for c in overall_d["원래업체"]],
                    ),
                    legend=alt.Legend(
                        orient="bottom", title=None,
                        columns=3, labelLimit=300, labelFontSize=14,
                    ),
                ),
            )

            labels = base.mark_text(
                radius=145, fontSize=13, fontWeight="bold",
            ).encode(
                text="라벨:N",
                color=alt.Color(
                    "업체:N",
                    scale=alt.Scale(
                        domain=overall_d["업체"].tolist(),
                        range=[COMPANY_COLORS[c]
                               for c in overall_d["원래업체"]],
                    ),
                    legend=None,
                ),
            )

            donut = (arc + labels).properties(width=420, height=420)

            _, c, _ = st.columns([1, 6, 1])
            with c:
                st.altair_chart(donut, use_container_width=False)

        # ===== 신규 공지사항 =====
        st.divider()
        st.subheader("📢 신규 공지사항")

        notice_date = st.date_input(
            "공지 날짜",
            value=datetime.date.today(),
            key="notice_date",
        )
        notice_key = notice_date.isoformat()

        existing_notice = fetch_notices(notice_key)
        cur_content = ""
        if not existing_notice.empty:
            cur_content = str(existing_notice.iloc[0].get("공지내용", "") or "")

        notice_content = st.text_area(
            "공지 내용",
            value=cur_content,
            height=100,
            key="notice_content",
        )

        if st.button("💾 공지사항 저장", key="save_notice"):
            try:
                if notice_content.strip():
                    sb.table("notices").upsert({
                        "날짜": notice_key,
                        "공지내용": notice_content.strip(),
                    }, on_conflict="날짜").execute()
                    invalidate_cache()
                    st.success(f"{notice_key} 공지사항을 저장했습니다.")
                    st.rerun()
                else:
                    st.warning("공지 내용이 비어 있습니다.")
            except Exception as e:
                st.error(f"저장 실패: {e}")

        # 미납현황
        st.divider()
        st.subheader("미납현황")

        sel_date = st.date_input("날짜", value=datetime.date.today(),
                                 key="memo_date")
        date_key = sel_date.isoformat()

        existing = fetch_memos(date_key)
        by_company = {r["업체"]: r for _, r in existing.iterrows()} \
                     if not existing.empty else {}

        table_df = pd.DataFrame({
            "업체명": COMPANIES,
            "미납여부": [bool(by_company.get(c, {}).get("미납여부", False))
                         for c in COMPANIES],
            "미납내용": [str(by_company.get(c, {}).get("미납내용", "") or "")
                         for c in COMPANIES],
        })

        edited = st.data_editor(
            table_df,
            key="memo_table",
            use_container_width=True,
            hide_index=True,
            column_config={
                "업체명": st.column_config.TextColumn("업체명", disabled=True),
                "미납여부": st.column_config.CheckboxColumn("미납"),
                "미납내용": st.column_config.TextColumn("미납내용"),
            },
        )

        common_action = ""
        if not existing.empty:
            common_action = str(existing.iloc[0].get("조치항목", "") or "")
        action = st.text_area("미납 조치항목", value=common_action)

        if st.button("💾 미납현황 저장", key="save_memo"):
            try:
                payload = []
                for _, r in edited.iterrows():
                    if not bool(r["미납여부"]):
                        continue
                    payload.append({
                        "날짜": date_key,
                        "업체": r["업체명"],
                        "미납여부": True,
                        "미납내용": str(r["미납내용"] or ""),
                        "조치항목": action,
                    })
                if not payload:
                    st.warning("미납 체크된 업체가 없습니다.")
                else:
                    sb.table("memos").upsert(payload,
                                             on_conflict="날짜,업체").execute()
                    invalidate_cache()
                    st.success(f"{date_key} 미납현황을 저장했습니다.")
                    st.rerun()
            except Exception as e:
                st.error(f"저장 실패: {e}")


# ----------------------------------------------------------------------
# 📁 파일 업로드
# ----------------------------------------------------------------------
with tab_upload:
    st.subheader("파일 업로드")
    st.caption("선택한 업체와 종류에 맞게 파일을 올리면 자동으로 파싱돼 저장됩니다.")

    up_company = st.selectbox("업체 선택", COMPANIES, key="up_company")
    up_kind = st.radio("파일 종류",
                       ["월별 (.xls/.xlsx)", "일별 (.csv)"],
                       horizontal=True, key="up_kind")

    allow_multi = (up_kind.startswith("일별") and up_company == "남양유업")
    if up_kind.startswith("일별"):
        if up_company == "남양유업":
            st.caption("※ 남양유업은 일별 파일 2개를 함께 올려주세요 (자동 합산).")
        else:
            st.caption("※ 일별 파일은 한 번에 하나씩 업로드합니다.")

    files = st.file_uploader(
        "파일을 끌어다 놓으세요" + (" (여러 개 가능)" if allow_multi else ""),
        type=["xls", "xlsx", "csv"],
        accept_multiple_files=allow_multi,
        key="up_files",
    )

    if files and st.button("🚀 업로드", key="up_btn"):
        file_list = files if isinstance(files, list) else [files]

        ok = fail = skip = 0
        progress = st.progress(0, text="업로드 중...")
        for idx, f in enumerate(file_list):
            raw = f.getvalue()
            filename = f.name
            try:
                kind_slug = "monthly" if up_kind.startswith("월별") else "daily"

                already = (sb.table("uploaded_files")
                             .select("id")
                             .eq("업체", up_company)
                             .eq("종류", kind_slug)
                             .eq("파일명", filename)
                             .execute().data)
                if already:
                    st.info(f"⏭ {filename}: 이미 처리된 파일입니다. 건너뜁니다.")
                    skip += 1
                    progress.progress((idx + 1) / len(file_list),
                                      text=f"{idx + 1}/{len(file_list)} 처리 중...")
                    continue

                path = (f"{company_code(up_company)}/{kind_slug}/"
                        f"{safe_path(filename)}")
                sb.storage.from_(BUCKET).upload(
                    path, raw,
                    file_options={"upsert": "true",
                                  "content-type": "application/octet-stream"},
                )

                if up_kind.startswith("월별"):
                    month, sales = parse_monthly(io.BytesIO(raw), filename, up_company)
                    if month is None or sales is None:
                        st.warning(f"⚠ {filename}: 월/판매금액을 못 찾았습니다.")
                        fail += 1
                    else:
                        sb.table("monthly").upsert({
                            "업체": up_company, "월": month, "판매금액": sales,
                        }, on_conflict="업체,월").execute()
                        sb.table("uploaded_files").insert({
                            "업체": up_company, "종류": kind_slug,
                            "파일명": filename, "월": month,
                        }).execute()
                        st.write(f"✅ {filename} → {month} / {sales:,}원")
                        ok += 1
                else:
                    date, revenue = parse_daily(io.BytesIO(raw), filename, up_company)
                    if date is None:
                        st.warning(f"⚠ {filename}: 파일명에서 날짜를 못 찾았습니다.")
                        fail += 1
                    else:
                        if up_company == "남양유업":
                            existing = (sb.table("daily")
                                          .select("매출액")
                                          .eq("업체", up_company)
                                          .eq("날짜", date)
                                          .execute().data)
                            if existing:
                                prev = int(existing[0]["매출액"])
                                new_total = prev + revenue
                                st.write(f"✅ {filename} → {date} / "
                                         f"기존 {prev:,} + 신규 {revenue:,} "
                                         f"= 합계 {new_total:,}원")
                                revenue = new_total
                            else:
                                st.write(f"✅ {filename} → {date} / {revenue:,}원")
                        else:
                            st.write(f"✅ {filename} → {date} / {revenue:,}원")

                        sb.table("daily").upsert({
                            "업체": up_company, "날짜": date, "매출액": revenue,
                        }, on_conflict="업체,날짜").execute()
                        sb.table("uploaded_files").insert({
                            "업체": up_company, "종류": kind_slug,
                            "파일명": filename, "날짜": date,
                        }).execute()
                        ok += 1
            except Exception as e:
                st.error(f"❌ {filename}: {e}")
                fail += 1
            progress.progress((idx + 1) / len(file_list),
                              text=f"{idx + 1}/{len(file_list)} 처리 중...")
        progress.empty()
        invalidate_cache()
        st.success(f"완료: 성공 {ok}건 · 실패 {fail}건 · 건너뜀 {skip}건")


# ----------------------------------------------------------------------
# 💰 매출 이력
# ----------------------------------------------------------------------
with tab_hist:
    st.subheader("매출 이력")
    st.caption("DB에 저장된 월별/일별 매출, 미납현황, 공지사항을 표로 조회합니다.")

    view = st.radio("조회 대상",
                    ["월별 매출 (monthly)", "일별 누적 매출 (daily)",
                     "미납현황 (memos)", "공지사항 (notices)"],
                    horizontal=True)

    if view.startswith("공지"):
        sel_company = "전체"
        c2, c3 = st.columns(2)
    else:
        c1, c2, c3 = st.columns(3)
        sel_company = c1.selectbox("업체", ["전체"] + COMPANIES,
                                    key="hist_company")

    if view.startswith("월별"):
        df = fetch_monthly()
        date_col = "월"
    elif view.startswith("일별"):
        df = fetch_daily()
        date_col = "날짜"
    elif view.startswith("미납"):
        df = fetch_memos()
        date_col = "날짜"
    else:
        df = fetch_notices()
        date_col = "날짜"

    if sel_company != "전체" and not df.empty and "업체" in df.columns:
        df = df[df["업체"] == sel_company]

    if not df.empty:
        all_dates = sorted(df[date_col].unique())
        if len(all_dates) >= 1:
            sel_from = c2.selectbox("시작", all_dates, index=0,
                                     key="hist_from")
            sel_to = c3.selectbox("종료", all_dates,
                                   index=len(all_dates) - 1,
                                   key="hist_to")
            df = df[(df[date_col] >= sel_from) & (df[date_col] <= sel_to)]

    sort_cols = [date_col, "업체"] if "업체" in df.columns else [date_col]
    df = df.sort_values(sort_cols).reset_index(drop=True)

    if "id" in df.columns:
        df = df.drop(columns=["id"])
    if "업로드시각" in df.columns:
        kst = datetime.timezone(datetime.timedelta(hours=9))
        df["업로드시각"] = (
            pd.to_datetime(df["업로드시각"], utc=True, errors="coerce")
              .dt.tz_convert(kst)
              .dt.strftime("%Y-%m-%d %H:%M:%S")
        )
    if "수정시각" in df.columns:
        kst = datetime.timezone(datetime.timedelta(hours=9))
        df["수정시각"] = (
            pd.to_datetime(df["수정시각"], utc=True, errors="coerce")
              .dt.tz_convert(kst)
              .dt.strftime("%Y-%m-%d %H:%M:%S")
        )

    if view.startswith("공지"):
        keep_cols = [c for c in ["날짜", "공지내용"] if c in df.columns]
        df = df[keep_cols]

    if df.empty:
        st.info("조회 결과가 없습니다.")
    else:
        st.write(f"총 {len(df)}건")
        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "판매금액": st.column_config.NumberColumn("판매금액", format="%,d"),
                "매출액": st.column_config.NumberColumn("매출액", format="%,d"),
            },
        )

        if view.startswith("월별") and "판매금액" in df.columns:
            st.caption(f"합계 (필터 범위): {int(df['판매금액'].sum()):,} 원")
        elif view.startswith("일별") and "매출액" in df.columns:
            st.caption(f"합계 (필터 범위): {int(df['매출액'].sum()):,} 원")

        st.download_button(
            "📥 CSV 다운로드",
            df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
            file_name=f"{view.split()[0]}_{datetime.date.today().isoformat()}.csv",
            mime="text/csv",
        )


# ----------------------------------------------------------------------
# 🎯 목표 설정 (LG생활건강 · 월별)
# ----------------------------------------------------------------------
with tab_target:
    st.subheader("목표 설정")
    st.caption("LG생활건강의 월별 목표 금액을 설정합니다. "
               "이전 목표는 자동으로 DB에 보존됩니다.")

    company = "LG생활건강"

    sel_target_month = st.text_input(
        "월 (YYYY-MM)",
        value=datetime.date.today().strftime("%Y-%m"),
        key="target_month",
    )

    cur = TARGETS.get(company, {}).get(sel_target_month, 0)
    new_target = st.number_input(
        f"{company} {sel_target_month} 목표 금액 (원)",
        min_value=0, value=int(cur), step=10_000_000,
        key="target_input",
    )
    st.caption(f"입력값: {int(new_target):,} 원")

    if st.button("💾 저장", key="save_target_new"):
        sb.table("targets").upsert({
            "업체": company,
            "월": sel_target_month,
            "목표금액": int(new_target),
        }, on_conflict="업체,월").execute()
        invalidate_cache()
        st.success(f"{company} {sel_target_month} 목표를 "
                   f"{int(new_target):,} 원으로 저장했습니다.")
        st.rerun()

    st.divider()
    st.subheader("저장된 목표 이력")
    rows = (sb.table("targets").select("*")
              .eq("업체", company)
              .order("월", desc=True).execute().data)
    if rows:
        target_df = pd.DataFrame(rows).sort_values("월", ascending=False)
        st.dataframe(
            target_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "목표금액": st.column_config.NumberColumn("목표금액", format="%,d"),
            },
        )
    else:
        st.info("저장된 목표가 없습니다.")
