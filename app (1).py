# -*- coding: utf-8 -*-
"""
app.py

유튜브 채널 영상 리포트 생성기 — Streamlit 웹앱 (YouTube Data API v3 버전).

로컬 실행:
    streamlit run app.py

API 키 설정 (Streamlit Cloud):
    앱 관리 화면 -> Settings -> Secrets 에 아래처럼 입력:
        youtube_api_key = "여기에_API_키"

API 키 설정 (Render):
    Dashboard -> 해당 서비스 -> Environment 탭에서 Environment Variable 추가:
        Key: YOUTUBE_API_KEY   Value: 여기에_API_키
    (st.secrets 파일이 없어도 동작하도록, 아래 get_api_key()가 환경변수도 함께 확인함)

API 키 설정 (로컬 테스트):
    .streamlit/secrets.toml 파일을 만들고 위 Streamlit Cloud와 동일하게 입력.
    또는 환경변수 YOUTUBE_API_KEY 를 셸에서 export 해도 됨.
"""

import io
import os
from datetime import date, datetime, timedelta

import streamlit as st

from core import (
    collect_multi_channel_videos,
    build_multi_channel_excel,
    TRANSLATOR_AVAILABLE,
    YouTubeAPIError,
)

st.set_page_config(
    page_title="유튜브 채널 영상 리포트",
    page_icon="🎬",
    layout="centered",
)

# ---------------------------------------------------------------------------
# 스타일
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700;800&display=swap');

    html, body, [class*="css"]  {
        font-family: 'Manrope', 'Pretendard', -apple-system, sans-serif;
    }

    .main {
        background: linear-gradient(180deg, #0f1220 0%, #14162a 100%);
    }

    .block-container {
        padding-top: 2.5rem;
        max-width: 760px;
    }

    .hero-title {
        font-size: 2.1rem;
        font-weight: 800;
        background: linear-gradient(90deg, #7C6CFF 0%, #FF6CAB 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.2rem;
    }
    .hero-sub {
        color: #9CA3AF;
        font-size: 0.95rem;
        margin-bottom: 0.4rem;
    }
    .credit {
        color: #6B7280;
        font-size: 0.78rem;
        margin-bottom: 1.6rem;
    }
    .hidden-credit {
        color: #6B7280;
        font-size: 0.7rem;
        text-align: right;
        margin-top: 2rem;
    }

    .card {
        background: rgba(255,255,255,0.04);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 16px;
        padding: 1.4rem 1.5rem;
        margin-bottom: 1.1rem;
        backdrop-filter: blur(6px);
    }
    .card-title {
        font-weight: 700;
        font-size: 1rem;
        color: #E5E7EB;
        margin-bottom: 0.7rem;
        display: flex;
        align-items: center;
        gap: 0.4rem;
    }

    .badge {
        display: inline-block;
        padding: 0.25rem 0.7rem;
        border-radius: 999px;
        font-size: 0.8rem;
        font-weight: 700;
        margin-right: 0.4rem;
    }
    .badge-long { background: rgba(89,140,255,0.18); color: #8FB4FF; }
    .badge-short { background: rgba(255,140,89,0.18); color: #FFB08F; }
    .badge-live { background: rgba(89,214,140,0.18); color: #8FE0B4; }
    .badge-total { background: rgba(124,108,255,0.22); color: #C4BAFF; }

    div.stButton > button, div.stFormSubmitButton > button, div.stDownloadButton > button {
        border-radius: 12px;
        font-weight: 700;
        border: none;
        padding: 0.6rem 1rem;
    }
    div.stFormSubmitButton > button[kind="primary"], div.stDownloadButton > button[kind="primary"] {
        background: linear-gradient(90deg, #7C6CFF 0%, #FF6CAB 100%);
        color: white;
    }

    div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div {
        border-radius: 10px !important;
    }

    .streamlit-expanderHeader {
        font-weight: 600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="hero-title">🎬 채널 영상 리포트</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="hero-sub">채널·기간·키워드를 입력하면, 정리된 엑셀 파일로 받아볼 수 있어요.</div>',
    unsafe_allow_html=True,
)
st.markdown('<div class="credit">제작자: 어둠의해커</div>', unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# API 키 확인
# ---------------------------------------------------------------------------
def get_api_key() -> str | None:
    """
    st.secrets(Streamlit Cloud 방식)와 환경변수(Render 등 방식) 둘 다 시도한다.
    secrets.toml 파일 자체가 없는 배포 환경(Render 등)에서는 st.secrets 접근이
    예외를 던질 수 있어 안전하게 감싼다.
    """
    try:
        value = st.secrets.get("youtube_api_key")
        if value:
            return value
    except Exception:
        pass

    return os.environ.get("YOUTUBE_API_KEY")


api_key = get_api_key()

if not api_key:
    st.error(
        "YouTube API 키가 설정되지 않았어요. "
        "앱 관리자에게 문의해주세요. (Secrets 또는 환경변수에 API 키 설정 필요)"
    )
    st.stop()

# ---------------------------------------------------------------------------
# 입력 폼
# ---------------------------------------------------------------------------
with st.form("report_form"):

    st.markdown('<div class="card"><div class="card-title">📺 채널</div>', unsafe_allow_html=True)
    channels_raw = st.text_area(
        "채널 주소 또는 @핸들 (여러 개면 줄바꿈으로 구분)",
        placeholder="@cloudair\nhttps://www.youtube.com/@rivalchannel\n@another_channel",
        height=90,
        label_visibility="collapsed",
    )
    st.caption("한 줄에 하나씩 입력하면 채널 여러 개를 한 번에 조회해서 엑셀 하나로 모아드려요.")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card"><div class="card-title">📅 기간</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        start_d = st.date_input("시작 날짜", value=date.today() - timedelta(days=30), format="YYYY-MM-DD")
    with col2:
        end_d = st.date_input("종료 날짜", value=date.today(), format="YYYY-MM-DD")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card"><div class="card-title">🔍 키워드 필터 (제목 기준)</div>', unsafe_allow_html=True)
    inc_col, exc_col = st.columns(2)
    with inc_col:
        include_raw = st.text_input(
            "필수 포함 키워드 (쉼표로 구분)",
            placeholder="예: 마케팅, 트렌드",
        )
        include_mode_label = st.radio(
            "포함 키워드 조건",
            ["하나라도 포함 (OR)", "모두 포함 (AND)"],
            horizontal=True,
        )
    with exc_col:
        exclude_raw = st.text_input(
            "제외 키워드 (쉼표로 구분)",
            placeholder="예: 브이로그, 광고",
        )
    st.caption("비워두면 해당 조건 없이 전체를 가져와요.")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card"><div class="card-title">⚙️ 옵션</div>', unsafe_allow_html=True)
    opt_col1, opt_col2 = st.columns(2)
    with opt_col1:
        sort_label = st.selectbox(
            "정렬 기준",
            ["최신순", "조회수순", "좋아요순", "댓글순"],
        )
    with opt_col2:
        translate = st.checkbox("설명 외국어 → 한국어 자동 번역", value=True)
    check_shorts = st.checkbox(
        "롱폼/숏폼 정확히 구분 (약간 느려짐)",
        value=True,
        help="유튜브에 실시간으로 확인해서 정확하게 구분해요. 끄면 훨씬 빠르지만 모든 영상이 '확인불가'로 표시돼요.",
    )
    if not TRANSLATOR_AVAILABLE:
        st.info("번역 기능이 아직 준비되지 않아, 이번 실행에서는 원문으로만 표시돼요.")
    st.markdown("</div>", unsafe_allow_html=True)

    submitted = st.form_submit_button("✨ 엑셀 만들기", use_container_width=True, type="primary")

# ---------------------------------------------------------------------------
# 실행
# ---------------------------------------------------------------------------
if submitted:
    channel_list = [c.strip() for c in channels_raw.splitlines() if c.strip()]
    if not channel_list:
        st.error("채널을 최소 1개 이상 입력해주세요.")
        st.stop()

    if start_d > end_d:
        st.error("시작 날짜가 종료 날짜보다 늦을 수 없어요.")
        st.stop()

    start_date = datetime.combine(start_d, datetime.min.time())
    end_date = datetime.combine(end_d, datetime.min.time())

    include_keywords = [k.strip() for k in include_raw.split(",") if k.strip()] if include_raw else []
    exclude_keywords = [k.strip() for k in exclude_raw.split(",") if k.strip()] if exclude_raw else []
    include_mode = "AND" if include_mode_label.startswith("모두") else "OR"

    sort_by_map = {"최신순": "date", "조회수순": "views", "좋아요순": "likes", "댓글순": "comments"}
    sort_by = sort_by_map[sort_label]

    status_box = st.empty()
    progress_bar = st.progress(0, text="시작하는 중...")
    log_box = st.expander("진행 상황 로그 보기", expanded=False)
    log_lines = []

    def progress_callback(msg: str):
        log_lines.append(msg)
        log_box.code("\n".join(log_lines[-300:]))

    channel_results = []
    total = len(channel_list)

    try:
        for i, raw in enumerate(channel_list):
            progress_bar.progress(
                int(i / total * 100),
                text=f"채널 조회 중... ({i+1}/{total}) {raw}",
            )
            status_box.info(f"'{raw}' 채널을 확인하고 있어요...")

            single_result = collect_multi_channel_videos(
                [raw], api_key, start_date, end_date,
                translate=translate, verbose=False, progress_callback=progress_callback,
                include_keywords=include_keywords, exclude_keywords=exclude_keywords,
                include_mode=include_mode, sort_by=sort_by, check_shorts=check_shorts,
            )
            channel_results.extend(single_result)

        progress_bar.progress(100, text="완료!")
    except Exception as e:
        progress_bar.empty()
        status_box.empty()
        st.error(f"영상 정보를 가져오는 중 문제가 발생했어요: {e}")
        st.caption("채널 주소가 정확한지 확인해보세요.")
        st.stop()

    status_box.empty()
    progress_bar.empty()

    all_videos = [v for cr in channel_results for v in cr["videos"]]

    if not all_videos:
        st.warning("조건에 맞는 영상을 찾지 못했어요. 채널 주소, 날짜 범위, 키워드 조건을 확인해주세요.")
        st.stop()

    long_count = len([v for v in all_videos if v["type"] == "롱폼"])
    short_count = len([v for v in all_videos if v["type"] == "숏폼"])
    live_count = len([v for v in all_videos if v["type"] == "라이브"])

    st.markdown(
        f"""
        <div class="card">
            <span class="badge badge-total">전체 {len(all_videos)}개</span>
            <span class="badge badge-long">롱폼 {long_count}개</span>
            <span class="badge badge-short">숏폼 {short_count}개</span>
            <span class="badge badge-live">라이브 {live_count}개</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    buffer = io.BytesIO()
    build_multi_channel_excel(channel_results, buffer)
    buffer.seek(0)

    first_name = channel_list[0].strip().replace("/", "_").replace(" ", "_").replace("@", "")
    suffix = f"외{len(channel_list)-1}개채널" if len(channel_list) > 1 else ""
    file_name = f"youtube_report_{first_name}{suffix}_{start_d.strftime('%Y%m%d')}_{end_d.strftime('%Y%m%d')}.xlsx"

    st.download_button(
        label="📥 엑셀 파일 다운로드",
        data=buffer,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

    for cr in channel_results:
        label = cr["channel_input"]
        videos = cr["videos"]
        with st.expander(f"📺 {label} — {len(videos)}개 영상", expanded=(len(channel_results) == 1)):
            if not videos:
                st.caption("이 채널에서는 조건에 맞는 영상을 찾지 못했어요.")
                continue
            st.dataframe(
                [
                    {
                        "제목": v["title"],
                        "날짜": v["upload_date"].strftime("%Y-%m-%d"),
                        "구분": v["type"],
                        "조회수": v.get("view_count") or "",
                        "좋아요": v.get("like_count") or "",
                        "댓글": v.get("comment_count") or "",
                    }
                    for v in videos
                ],
                use_container_width=True,
                hide_index=True,
            )

st.divider()
with st.expander("ℹ️ 사용 팁 / 참고 사항"):
    st.markdown(
        """
- 이제 유튜브 공식 API를 사용해서, 이전보다 훨씬 안정적으로 동작해요.
- 롱폼/숏폼 구분은 유튜브에 실시간으로 물어봐서 정확하게 나눠요 (재생시간 추측 아님).
- 라이브 방송(진행중/예정/종료)은 공식 데이터로 100% 정확하게 구분돼요.
- 채널 여러 개를 한 번에 조회하면 시간이 그만큼 더 걸려요.
- 조회수·좋아요·댓글수는 **수집한 시점 기준**으로 고정돼요.
- '내용 요약'은 AI가 새로 쓴 게 아니라, 영상 설명 앞부분(최대 200자)을 가져와 필요시 한국어로 번역한 거예요.
- 키워드 필터는 **제목 기준**으로만 검사해요.
- 비공개 채널이거나 업로드 목록을 막아둔 채널은 조회가 안 될 수 있어요.
        """
    )

st.markdown('<div class="hidden-credit">for 돌맹이</div>', unsafe_allow_html=True)
