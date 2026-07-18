import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import json
from pathlib import Path
from agent import CFPSAgent, file_to_text, CATEGORIES
from cfps_parser import parse_cfps_excel

st.set_page_config(
    page_title="CFPS Cost Calculator",
    page_icon=":material/science:",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════════════════════
# 디자인 시스템 토큰
# Source: IBM Carbon (colors, spacing, type) + Radix UI (teal scale)
#         + Ant Design (tab, table patterns)
# ══════════════════════════════════════════════════════════════════════════════

# Carbon Data Visualization 공식 14색 팔레트 → 카테고리 매핑
CATEGORY_COLORS = {
    "추출물":   "#da1e28",  # Carbon Red-60
    "에너지":   "#b28600",  # Carbon Yellow-60
    "NTPs":    "#0072c3",  # Carbon Blue-60
    "아미노산": "#198038",  # Carbon Green-60
    "버퍼/염류":"#007d79",  # Carbon Teal-60
    "보조인자": "#6929c4",  # Carbon Purple-60
    "tRNA":    "#9f1853",  # Carbon Magenta-60
    "첨가물":  "#525252",  # Carbon Gray-60
    "DNA 주형":"#005d5d",  # Carbon Teal-70
    "기타":    "#8d8d8d",  # Carbon Gray-50
}

# Plotly 공통 레이아웃 (Carbon simple-white 스타일)
CHART_LAYOUT = dict(
    template="simple_white",
    font=dict(family="'IBM Plex Sans', sans-serif", size=12, color="#161616"),
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    margin=dict(t=48, b=0, l=0, r=0),
    colorway=list(CATEGORY_COLORS.values()),
)

# ── Global CSS (IBM Plex Sans + Carbon/Ant Design 컴포넌트 스타일) ─────────────
st.html("""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:ital,wght@0,300;0,400;0,600;1,400&family=IBM+Plex+Mono:wght@400&family=Noto+Sans+KR:wght@300;400;600&display=swap" rel="stylesheet">
<style>

/* ── 전체 폰트: IBM Plex Sans (Latin) + Noto Sans KR (한글 폴백) ── */
*, *::before, *::after {
    font-family: 'IBM Plex Sans', 'Noto Sans KR', -apple-system, BlinkMacSystemFont, sans-serif !important;
}
code, pre, [data-testid="stCode"] * {
    font-family: 'IBM Plex Mono', 'Courier New', monospace !important;
}

/* ── 제목 (letter-spacing 제거 — 한글에 부적합) ── */
h1 { font-weight: 300 !important; font-size: 1.9rem !important; line-height: 1.25 !important; }
h2 { font-weight: 600 !important; font-size: 1.2rem !important; }
h3 { font-weight: 600 !important; font-size: 1rem !important; }

/* ── KPI 타일 (Carbon: 상단 포인트 선 + 흰 배경) ── */
[data-testid="metric-container"] {
    background: #ffffff !important;
    border: 1px solid #e0e0e0 !important;
    border-top: 3px solid #12a594 !important;
    border-radius: 0 !important;
    padding: 1rem 1.25rem 1.25rem !important;
}
[data-testid="stMetricLabel"] > div {
    font-size: 12px !important;
    font-weight: 600 !important;
    color: #525252 !important;
    margin-bottom: 6px !important;
    /* uppercase·letter-spacing 제거 — 한글 깨짐 원인 */
}
[data-testid="stMetricValue"] > div {
    font-size: 1.6rem !important;
    font-weight: 400 !important;
    color: #161616 !important;
    line-height: 1.1 !important;
}
[data-testid="stMetricDelta"] > div {
    font-size: 12px !important;
    font-weight: 400 !important;
}

/* ── 탭 (Ant Design underline 스타일) ── */
.stTabs [data-baseweb="tab-list"] {
    background: transparent !important;
    border-bottom: 1px solid #e0e0e0 !important;
    gap: 0 !important;
    padding: 0 !important;
    margin-bottom: 0 !important;
}
.stTabs [data-baseweb="tab"] {
    border-bottom: 2px solid transparent !important;
    border-radius: 0 !important;
    padding: 12px 24px !important;
    font-size: 13px !important;
    font-weight: 400 !important;
    color: #525252 !important;
    background: transparent !important;
    margin-bottom: -1px !important;
}
.stTabs [aria-selected="true"][data-baseweb="tab"] {
    border-bottom: 2px solid #12a594 !important;
    color: #12a594 !important;
    font-weight: 600 !important;
}

/* ── 카드 컨테이너 (Carbon: 0 radius, 밝은 테두리) ── */
[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 0 !important;
    border-color: #e0e0e0 !important;
    background: #ffffff !important;
}

/* ── 버튼 ── */
button[kind="primary"], [data-testid="baseButton-primary"] {
    border-radius: 0 !important;
    font-weight: 500 !important;
    font-size: 14px !important;
}
[data-testid="baseButton-secondary"] {
    border-radius: 0 !important;
    font-weight: 400 !important;
    font-size: 14px !important;
}

/* ── 입력 ── */
input, textarea, [data-baseweb="input"] {
    border-radius: 0 !important;
}
[data-baseweb="select"] > div {
    border-radius: 0 !important;
}

/* ── 사이드바 ── */
[data-testid="stSidebar"] {
    border-right: 1px solid #e0e0e0 !important;
}

/* ── 프로그레스 바 ── */
[data-testid="stProgressBar"] > div,
[data-testid="stProgressBar"] > div > div {
    border-radius: 0 !important;
}

/* ── Expander ── */
[data-testid="stExpander"] {
    border-radius: 0 !important;
    border-color: #e0e0e0 !important;
}

/* ── 채팅 메시지 ── */
[data-testid="stChatMessage"] {
    border-radius: 0 !important;
    background: #f4f4f4 !important;
    border: 1px solid #e0e0e0 !important;
}

/* ── 알림 박스 ── */
[data-testid="stAlert"] {
    border-radius: 0 !important;
}

/* ── 파일 업로더 ── */
[data-testid="stFileUploaderDropzone"] {
    border-radius: 0 !important;
    background: #f4f4f4 !important;
    border: 1px dashed #8d8d8d !important;
}

/* ── Pills ── */
[data-testid="stPills"] button {
    border-radius: 20px !important;
    font-size: 12px !important;
}

</style>
""")

# ── 데이터 ──────────────────────────────────────────────────────────────────
PRESET_FILE = Path(__file__).parent / "reagents_preset.json"

DEFAULT_REAGENTS = [
    ("E. coli 추출물 (extract)", "추출물",   3.00, 1.0,  50000),
    ("PEP (phosphoenolpyruvate)","에너지",    0.50, 1.0,  15000),
    ("ATP",                      "NTPs",     0.20, 1.0,   8000),
    ("GTP",                      "NTPs",     0.20, 1.0,   8000),
    ("CTP",                      "NTPs",     0.20, 1.0,   8000),
    ("UTP",                      "NTPs",     0.20, 1.0,   5000),
    ("아미노산 혼합 (20종)",       "아미노산",  1.00, 1.0,  10000),
    ("HEPES buffer",             "버퍼/염류", 0.30, 1.0,   3000),
    ("Mg-glutamate",             "버퍼/염류", 0.20, 1.0,   2000),
    ("K-glutamate",              "버퍼/염류", 0.30, 1.0,   2000),
    ("NAD",                      "보조인자",  0.10, 1.0,   5000),
    ("CoA",                      "보조인자",  0.10, 1.0,   8000),
    ("Spermidine",               "보조인자",  0.05, 1.0,   2000),
    ("Putrescine",               "보조인자",  0.05, 1.0,   1000),
    ("cAMP",                     "보조인자",  0.05, 1.0,   3000),
    ("Folinic acid",             "보조인자",  0.05, 1.0,   5000),
    ("tRNA (E. coli)",           "tRNA",     0.20, 0.5,  30000),
    ("PEG-8000",                 "첨가물",   0.50, 5.0,   5000),
    ("플라스미드 DNA",             "DNA 주형", 0.50, 0.1,   5000),
    ("RNase inhibitor",          "기타",     0.20, 0.5,  20000),
]

EXAMPLE_PROMPTS = [
    "extract 단가를 mL당 60,000원으로 바꿔줘",
    "ATP 사용량을 0.3 µL로 수정해줘",
    "creatine phosphate 0.5 µL 추가해줘",
    "RNase inhibitor 제외해줘",
]


def make_default_df() -> pd.DataFrame:
    rows = []
    for name, cat, vol, ml, cost in DEFAULT_REAGENTS:
        rows.append({
            "포함": True, "성분": name, "카테고리": cat,
            "반응당 사용량 (µL)": vol, "시약 총량 (mL)": ml,
            "시약 총 비용 (₩)": cost,
            "반응당 비용 (₩)": round(vol / (ml * 1000) * cost, 2),
        })
    return pd.DataFrame(rows)


def recalc(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["반응당 비용 (₩)"] = (
        df["반응당 사용량 (µL)"] / (df["시약 총량 (mL)"] * 1000) * df["시약 총 비용 (₩)"]
    ).round(2)
    return df


def records_to_df(records: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(records)
    cols = ["포함","성분","카테고리","반응당 사용량 (µL)","시약 총량 (mL)","시약 총 비용 (₩)","반응당 비용 (₩)"]
    for c in cols:
        if c not in df.columns:
            df[c] = True if c == "포함" else 0.0
    return recalc(df[cols])


def load_preset():
    if PRESET_FILE.exists():
        with open(PRESET_FILE, encoding="utf-8") as f:
            return records_to_df(json.load(f))
    return None


def save_preset(df: pd.DataFrame):
    with open(PRESET_FILE, "w", encoding="utf-8") as f:
        json.dump(df.to_dict(orient="records"), f, ensure_ascii=False, indent=2)


# ── 세션 ──────────────────────────────────────────────────────────────────────
if "reagents_df"  not in st.session_state:
    st.session_state["reagents_df"]  = load_preset() or make_default_df()
if "chat_history" not in st.session_state:
    st.session_state["chat_history"] = []


# ══════════════════════════════════════════════════════════════════════════════
# 사이드바
# ══════════════════════════════════════════════════════════════════════════════
def section_label(text: str):
    """사이드바·탭 공용 섹션 레이블 — uppercase/letter-spacing 없음(한글 안전)"""
    st.markdown(
        f'<p style="font-size:12px;font-weight:600;color:#525252;margin:0 0 8px;">{text}</p>',
        unsafe_allow_html=True,
    )


with st.sidebar:
    # 브랜드 헤더 — uppercase는 영문 부제목에만 사용
    st.html("""
    <div style="padding:1rem 0; border-bottom:1px solid #e0e0e0; margin-bottom:1rem;">
      <div style="font-size:10px; font-weight:600; text-transform:uppercase;
                  letter-spacing:0.1em; color:#525252; margin-bottom:4px;">
        CFPS · Cell-Free System
      </div>
      <div style="font-size:1.05rem; font-weight:600; color:#161616; line-height:1.4;">
        반응당 단가 계산기
      </div>
      <div style="font-size:12px; color:#525252; margin-top:3px;">
        E. coli 세포추출물 기반
      </div>
    </div>
    """)

    # 반응 설정
    section_label("반응 설정")
    with st.container(border=True):
        total_rxn_vol = st.number_input(
            "총 반응 부피 (µL)", min_value=1, max_value=500, value=10, step=1
        )
        num_rxn = st.number_input(
            "반응 수", min_value=1, max_value=10000, value=1, step=1
        )

    st.write("")

    # API Key
    section_label("Claude API")
    with st.container(border=True):
        api_key = st.text_input(
            "Anthropic API Key",
            type="password",
            placeholder="sk-ant-...",
            label_visibility="collapsed",
        )
        if api_key:
            st.success("연결됨", icon=":material/check_circle:")
        else:
            st.caption(":material/info: AI 에이전트 사용 시 필요")

    st.write("")

    # 수율 설정
    section_label("수율 기반 단가")
    with st.container(border=True):
        show_yield = st.toggle("단백질 수율 입력", value=False)
        protein_yield = None
        if show_yield:
            protein_yield = st.number_input(
                "수율 (µg/mL)", min_value=0.0, value=100.0, step=10.0
            )

    st.write("")

    c1, c2 = st.columns(2)
    if c1.button(":material/save: 저장", use_container_width=True):
        save_preset(st.session_state["reagents_df"])
        st.toast("저장 완료", icon=":material/check:")
    if c2.button(":material/refresh: 초기화", use_container_width=True):
        if PRESET_FILE.exists():
            PRESET_FILE.unlink()
        st.session_state.update({"reagents_df": make_default_df(), "chat_history": []})
        st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# 페이지 헤더 + KPI 스트립
# ══════════════════════════════════════════════════════════════════════════════
current_df = st.session_state["reagents_df"]
active_df  = current_df[current_df["포함"]].copy()
total_cost = active_df["반응당 비용 (₩)"].sum()
active_vol = active_df["반응당 사용량 (µL)"].sum()
n_active   = len(active_df)
n_total    = len(current_df)

# 페이지 제목 — uppercase는 영문에만, 한글 제목은 letter-spacing 없음
st.html("""
<div style="padding:1rem 0 1.25rem; border-bottom:1px solid #e0e0e0; margin-bottom:1.25rem;">
  <div style="font-size:10px; font-weight:600; text-transform:uppercase;
              letter-spacing:0.1em; color:#525252; margin-bottom:6px;">
    Synthetic Biology · Cost Analysis
  </div>
  <div style="font-size:1.65rem; font-weight:300; color:#161616; line-height:1.25;">
    CFPS 반응당 단가 계산기
  </div>
  <div style="font-size:13px; color:#525252; margin-top:5px;">
    E. coli 세포추출물 기반 · Claude AI 에이전트 지원
  </div>
</div>
""")

# KPI 타일 (4개 — Carbon white tile + teal top border)
k1, k2, k3, k4 = st.columns(4)
k1.metric("반응당 총 비용",     f"₩ {total_cost:,.1f}")
k2.metric(f"{num_rxn}회 총 비용", f"₩ {total_cost * num_rxn:,.0f}", f"{num_rxn}회 기준")
k3.metric("µL당 비용",          f"₩ {total_cost/total_rxn_vol:.2f}" if total_rxn_vol else "—")
k4.metric("포함 시약",          f"{n_active} / {n_total}종")

st.write("")


# ══════════════════════════════════════════════════════════════════════════════
# 탭 (Ant Design underline 스타일 — CSS에서 적용)
# ══════════════════════════════════════════════════════════════════════════════
tab_agent, tab_table, tab_chart, tab_summary = st.tabs(
    ["AI 에이전트", "시약 목록", "비용 분석", "요약"]
)


# ── 탭 1: AI 에이전트 ──────────────────────────────────────────────────────────
with tab_agent:
    st.write("")

    # ── 파일 업로드 섹션 ─────────────────────────────────────────────────
    section_label("1 · 파일 업로드로 자동 파싱")

    with st.container(border=True):
        up_left, up_right = st.columns([3, 2])
        with up_left:
            uploaded = st.file_uploader(
                "파일 선택",
                type=["xlsx","xls","csv","pdf","txt","md","docx"],
                label_visibility="collapsed",
                help="CFPS 조성 Excel, 구매 영수증, 프로토콜 파일 지원",
            )
        with up_right:
            st.html("""
            <div style="font-size:12px; color:#525252; line-height:1.9; padding:4px 0;">
              <strong style="color:#161616;">지원 형식</strong><br>
              <span style="color:#12a594;">★</span> CFPS 조성 Excel — 전용 파서 자동 적용<br>
              Excel / CSV — 구매 목록, 재고 관리표<br>
              PDF — 견적서, 영수증<br>
              Word / 텍스트 — 실험 프로토콜
            </div>
            """)

        if uploaded:
            st.caption(f":material/attach_file: **{uploaded.name}** · {uploaded.size:,} bytes")

            # ── CFPS 전용 파서 자동 감지 ─────────────────────────────────
            is_cfps_format = uploaded.name.lower().endswith((".xlsx", ".xls"))
            cfps_detected  = False
            if is_cfps_format:
                try:
                    import io, pandas as _pd
                    _raw = uploaded.read()
                    uploaded.seek(0)
                    _xl = _pd.ExcelFile(io.BytesIO(_raw))
                    _sheets = [s.lower() for s in _xl.sheet_names]
                    cfps_detected = any("reaction buffer" in s or "cell-free condition" in s for s in _sheets)
                except Exception:
                    pass

            if cfps_detected:
                st.info(
                    "**CFPS 조성 Excel 감지됨** — 전용 파서로 자동 분석합니다. "
                    "반응 볼륨과 성분별 사용량을 직접 계산합니다.",
                    icon=":material/auto_awesome:",
                )
                parse_btn = st.button(
                    ":material/science: CFPS 파일 파싱", type="primary"
                )
                if parse_btn:
                    with st.spinner("반응 조성 분석 중..."):
                        uploaded.seek(0)
                        result = parse_cfps_excel(uploaded)
                    if result["errors"]:
                        for e in result["errors"]:
                            st.warning(e, icon=":material/warning:")
                    if result["reagents"]:
                        new_df = pd.DataFrame(result["reagents"])
                        # 기존 컬럼 중 누락된 것 채우기
                        for col in ["가격 출처 URL","가격 조회일","가격 메모","공급사","CAS No."]:
                            if col not in new_df.columns:
                                new_df[col] = ""
                        for col in ["시약 총량 (mL)","시약 총 비용 (₩)","반응당 비용 (₩)"]:
                            if col not in new_df.columns:
                                new_df[col] = 0.0
                        st.session_state["reagents_df"] = new_df
                        st.session_state["rxn_vol_override"] = result["total_rxn_vol"]
                        st.success(
                            f"{len(result['reagents'])}개 시약 추출 완료 "
                            f"(반응 부피 {result['total_rxn_vol']} µL) — "
                            "가격이 없는 시약은 아래 '가격 자동 조회'를 사용하세요.",
                            icon=":material/check_circle:",
                        )
                        st.rerun()
            else:
                # 기존 Claude 에이전트 파싱
                btn_col, _ = st.columns([1, 3])
                parse_btn = btn_col.button(
                    ":material/search: AI 분석 시작", type="primary",
                    use_container_width=True, disabled=not api_key,
                )
                if not api_key:
                    st.warning("사이드바에서 API Key를 먼저 입력하세요.", icon=":material/key:")
                elif parse_btn:
                    with st.spinner(f"'{uploaded.name}' 분석 중..."):
                        file_text = file_to_text(uploaded)
                        if file_text.startswith("[") and "오류" in file_text:
                            st.error(file_text, icon=":material/error:")
                        else:
                            agent   = CFPSAgent(api_key)
                            records, msg = agent.parse_file(file_text, uploaded.name)
                            if records is not None:
                                st.session_state["reagents_df"] = records_to_df(records)
                                st.session_state["chat_history"].append(
                                    {"role": "assistant", "content": f"**{uploaded.name}** 파싱 완료\n\n{msg}"}
                                )
                                st.success(f"{len(records)}개 시약 추출 완료 — '시약 목록' 탭 확인", icon=":material/check_circle:")
                                st.rerun()
                            else:
                                st.error(f"파싱 실패: {msg}", icon=":material/error:")

    st.write("")
    st.write("")

    # ── 채팅 섹션 ────────────────────────────────────────────────────────
    section_label("2 · 채팅으로 테이블 수정")

    with st.container(border=True):
        # 예시 프롬프트 (Radix-style pill chips)
        selected = st.pills(
            "예시 명령",
            options=EXAMPLE_PROMPTS,
            selection_mode="single",
            help="클릭하면 해당 명령이 실행됩니다.",
        )

        # 채팅 기록
        chat_box = st.container(height=280)
        with chat_box:
            if not st.session_state["chat_history"]:
                st.html("""
                <div style="height:200px; display:flex; align-items:center; justify-content:center;
                            flex-direction:column; gap:8px; color:#a8a8a8;">
                  <span style="font-size:32px;">💬</span>
                  <span style="font-size:13px;">파일을 분석하거나 아래 명령을 입력하세요</span>
                </div>
                """)
            else:
                for msg in st.session_state["chat_history"]:
                    with st.chat_message(msg["role"]):
                        st.markdown(msg["content"])

        user_input = st.chat_input(
            "시약 수정 명령 입력... (예: extract 단가를 60,000원으로 바꿔줘)",
            disabled=not api_key,
        ) or (selected if selected else None)

        if user_input and api_key:
            st.session_state["chat_history"].append({"role": "user", "content": user_input})
            with st.spinner("처리 중..."):
                agent   = CFPSAgent(api_key)
                records, reply = agent.chat(user_input, st.session_state["reagents_df"])
                if records is not None:
                    st.session_state["reagents_df"] = records_to_df(records)
                    reply = reply or "테이블을 업데이트했습니다."
                st.session_state["chat_history"].append({"role": "assistant", "content": reply})
            st.rerun()

        if st.session_state["chat_history"]:
            if st.button(":material/delete_sweep: 대화 초기화"):
                st.session_state["chat_history"] = []
                st.rerun()

    st.write("")

    # ── 가격 자동 조회 섹션 ───────────────────────────────────────────────
    section_label("3 · 가격 자동 조회 (웹 검색)")

    _no_price_df = st.session_state["reagents_df"]
    _no_price_n  = int((_no_price_df["시약 총 비용 (₩)"] == 0).sum())

    with st.container(border=True):
        left_p, right_p = st.columns([3, 2])
        with left_p:
            if _no_price_n > 0:
                st.markdown(
                    f"가격 미입력 시약 **{_no_price_n}종** — "
                    "웹 검색으로 Sigma-Aldrich·Thermo 등 현재 시장가를 자동 조회합니다."
                )
                st.caption("조회된 가격은 URL과 날짜와 함께 기록됩니다.")
            else:
                st.success("모든 시약에 가격이 입력되어 있습니다.", icon=":material/check:")
        with right_p:
            price_btn = st.button(
                ":material/travel_explore: 가격 자동 조회",
                type="primary",
                disabled=not api_key or _no_price_n == 0,
                use_container_width=True,
            )
            if not api_key:
                st.caption("API Key 필요")

        if price_btn and api_key:
            progress_area = st.empty()
            def _cb(msg):
                progress_area.info(msg, icon=":material/search:")

            with st.spinner("가격 조회 중 (웹 검색)..."):
                agent = CFPSAgent(api_key)
                updated_df, msg = agent.lookup_prices(
                    st.session_state["reagents_df"], callback=_cb
                )
            progress_area.empty()
            st.session_state["reagents_df"] = updated_df
            st.success(msg[:300], icon=":material/check_circle:")
            st.rerun()

        # 가격 이력 테이블 (URL + 날짜 있는 항목만)
        _price_logged = st.session_state["reagents_df"]
        if "가격 출처 URL" in _price_logged.columns:
            _has_url = _price_logged[_price_logged["가격 출처 URL"].str.len() > 0]
            if not _has_url.empty:
                st.write("")
                section_label("가격 이력 기록")
                _disp = _has_url[["성분","시약 총 비용 (₩)","시약 총량 (mL)","가격 출처 URL","가격 조회일"]].copy()
                _disp["시약 총 비용 (₩)"] = _disp["시약 총 비용 (₩)"].apply(lambda x: f"₩ {x:,.0f}")
                st.dataframe(
                    _disp,
                    column_config={
                        "가격 출처 URL": st.column_config.LinkColumn("출처 URL"),
                    },
                    use_container_width=True,
                    hide_index=True,
                )


# ── 탭 2: 시약 목록 ────────────────────────────────────────────────────────────
with tab_table:
    st.write("")

    # 부피 현황 바 (Carbon progress bar 스타일)
    active_vol = st.session_state["reagents_df"][st.session_state["reagents_df"]["포함"]]["반응당 사용량 (µL)"].sum()
    water_vol  = max(0.0, total_rxn_vol - active_vol)
    vol_pct    = min(active_vol / total_rxn_vol, 1.0) if total_rxn_vol else 0

    bar_col, stat_col = st.columns([3, 1])
    with bar_col:
        section_label("반응 부피 구성")
        st.markdown(
            f"시약 **{active_vol:.1f} µL** + H₂O **{water_vol:.1f} µL** = {total_rxn_vol} µL"
        )
        st.progress(vol_pct)
    with stat_col:
        st.metric("포함 시약", f"{n_active} / {n_total}종")

    st.write("")

    # 시약 추가
    with st.expander(":material/add: 시약 추가"):
        fc1, fc2, fc3, fc4, fc5, fc6 = st.columns([3,2,1.5,1.5,1.8,1])
        new_name = fc1.text_input("성분명", key="new_name", placeholder="예: Creatine phosphate")
        new_cat  = fc2.selectbox("카테고리", CATEGORIES, key="new_cat")
        new_vol  = fc3.number_input("사용량 (µL)", min_value=0.0, value=0.5, step=0.05, key="new_vol")
        new_ml   = fc4.number_input("총량 (mL)",   min_value=0.001, value=1.0, step=0.1, key="new_ml")
        new_cost = fc5.number_input("총 비용 (₩)", min_value=0, value=10000, step=1000, key="new_cost")
        fc6.write("")
        fc6.write("")
        if fc6.button(":material/add:", type="primary", use_container_width=True) and new_name:
            new_row = {
                "포함": True, "성분": new_name, "카테고리": new_cat,
                "반응당 사용량 (µL)": new_vol, "시약 총량 (mL)": new_ml,
                "시약 총 비용 (₩)": new_cost,
                "반응당 비용 (₩)": round(new_vol / (new_ml * 1000) * new_cost, 2),
            }
            st.session_state["reagents_df"] = pd.concat(
                [st.session_state["reagents_df"], pd.DataFrame([new_row])], ignore_index=True
            )
            st.rerun()

    # 테이블 (Ant Design 스타일 — 헤더 강조, 줄무늬)
    edited_df = st.data_editor(
        st.session_state["reagents_df"],
        column_config={
            "포함":              st.column_config.CheckboxColumn("포함", width="small"),
            "성분":              st.column_config.TextColumn("성분", width="large"),
            "카테고리":          st.column_config.SelectboxColumn("카테고리", options=CATEGORIES, width="medium"),
            "반응당 사용량 (µL)": st.column_config.NumberColumn("사용량 (µL/rxn)", min_value=0.0, format="%.3f"),
            "시약 총량 (mL)":    st.column_config.NumberColumn("시약 총량 (mL)", min_value=0.001, format="%.3f"),
            "시약 총 비용 (₩)":  st.column_config.NumberColumn("총 비용 (₩)", min_value=0, format="%d"),
            "반응당 비용 (₩)":   st.column_config.NumberColumn("반응당 비용 (₩)", disabled=True, format="%.2f"),
        },
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key="editor",
    )
    st.session_state["reagents_df"] = recalc(edited_df)


# ── 비용 집계 재계산 ──────────────────────────────────────────────────────────
current_df = st.session_state["reagents_df"]
active_df  = current_df[current_df["포함"]].copy()
total_cost = active_df["반응당 비용 (₩)"].sum()


# ── 탭 3: 비용 분석 ────────────────────────────────────────────────────────────
with tab_chart:
    st.write("")

    if active_df.empty:
        st.info("포함된 시약이 없습니다. '시약 목록' 탭에서 포함 여부를 체크하세요.", icon=":material/info:")
    else:
        cat_cost = (
            active_df.groupby("카테고리")["반응당 비용 (₩)"].sum()
            .reset_index()
            .rename(columns={"반응당 비용 (₩)": "비용"})
            .sort_values("비용", ascending=False)
        )
        cat_cost["비율 (%)"] = (cat_cost["비용"] / total_cost * 100).round(1)

        ch1, ch2 = st.columns(2)

        # 도넛 차트
        with ch1:
            section_label("카테고리별 비율")
            fig_donut = go.Figure(go.Pie(
                labels=cat_cost["카테고리"],
                values=cat_cost["비용"],
                hole=0.48,
                marker=dict(colors=[CATEGORY_COLORS.get(c,"#999") for c in cat_cost["카테고리"]],
                            line=dict(color="#f4f4f4", width=2)),
                textinfo="percent+label",
                textfont=dict(size=11),
                hovertemplate="%{label}<br>₩ %{value:,.1f} (%{percent})<extra></extra>",
            ))
            fig_donut.update_layout(
                showlegend=False,
                height=340,
                **CHART_LAYOUT,
                annotations=[dict(
                    text=f"₩{total_cost:,.0f}",
                    x=0.5, y=0.5, showarrow=False,
                    font=dict(size=14, color="#161616", family="IBM Plex Sans"),
                )]
            )
            st.plotly_chart(fig_donut, use_container_width=True)

        # 수평 바 차트
        with ch2:
            section_label("성분별 반응당 비용")
            comp_df = active_df[["성분","카테고리","반응당 비용 (₩)"]].sort_values("반응당 비용 (₩)", ascending=True)
            fig_bar = px.bar(
                comp_df, x="반응당 비용 (₩)", y="성분",
                color="카테고리", orientation="h",
                color_discrete_map=CATEGORY_COLORS,
            )
            fig_bar.update_layout(
                showlegend=False,
                height=max(320, len(comp_df) * 22),
                xaxis=dict(title="반응당 비용 (₩)", showgrid=True, gridcolor="#e0e0e0"),
                yaxis=dict(title=None),
                **CHART_LAYOUT,
            )
            fig_bar.update_traces(hovertemplate="%{y}<br>₩ %{x:,.2f}<extra></extra>")
            st.plotly_chart(fig_bar, use_container_width=True)

        # 카테고리 집계 테이블
        section_label("카테고리 집계")
        st.dataframe(
            cat_cost.rename(columns={"비용": "비용 (₩)"})
            .style
            .format({"비용 (₩)": "{:,.1f}", "비율 (%)": "{:.1f}%"})
            .bar(subset=["비용 (₩)"], color="#ccf3ea", vmin=0),
            use_container_width=True,
            hide_index=True,
        )


# ── 탭 4: 요약 ────────────────────────────────────────────────────────────────
with tab_summary:
    st.write("")

    # 비용 요약 타일
    section_label("비용 요약")
    with st.container(border=True):
        m1, m2, m3 = st.columns(3)
        m1.metric("반응당 총 비용",         f"₩ {total_cost:,.1f}")
        m2.metric(f"{num_rxn}회 반응 총 비용", f"₩ {total_cost * num_rxn:,.0f}")
        m3.metric("반응 부피당 비용",        f"₩ {total_cost/total_rxn_vol:.2f} / µL" if total_rxn_vol else "—")

    if protein_yield and protein_yield > 0:
        st.write("")
        section_label("단백질 수율 기반 단가")
        with st.container(border=True):
            protein_ug = protein_yield / 1000 * total_rxn_vol
            m4, m5, _ = st.columns(3)
            m4.metric("반응당 예상 수율",      f"{protein_ug:.2f} µg")
            m5.metric("단백질 1 µg당 비용",    f"₩ {total_cost/protein_ug:.2f}" if protein_ug > 0 else "—")

    st.write("")
    section_label("전체 성분 상세")
    with st.container(border=True):
        export_cols = ["성분","카테고리","반응당 사용량 (µL)","반응당 비용 (₩)"]
        for extra in ["공급사","CAS No.","시약 총 비용 (₩)","시약 총량 (mL)","가격 출처 URL","가격 조회일","가격 메모"]:
            if extra in active_df.columns:
                export_cols.append(extra)
        display_df = active_df[export_cols].copy()
        display_df["비율 (%)"] = (display_df["반응당 비용 (₩)"] / total_cost * 100).round(1)
        st.dataframe(
            display_df[["성분","카테고리","반응당 사용량 (µL)","반응당 비용 (₩)","비율 (%)"]].style
            .format({"반응당 사용량 (µL)": "{:.3f}","반응당 비용 (₩)": "{:,.2f}","비율 (%)": "{:.1f}%"})
            .bar(subset=["비율 (%)"], color="#ccf3ea", vmin=0, vmax=100),
            use_container_width=True,
            hide_index=True,
        )
        csv = display_df.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            ":material/download: CSV 다운로드",
            data=csv.encode("utf-8-sig"),
            file_name="cfps_cost_per_rxn.csv",
            mime="text/csv",
        )
