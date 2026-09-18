import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# 1. 페이지 레이아웃 및 브라우저 탭 설정
st.set_page_config(
    page_title="FITI 상해지사 경영 실적 분석 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. 고급 카드 및 폰트 디자인 커스텀 CSS
st.markdown("""
<style>
    /* 전체 폰트 및 배경 정돈 */
    .block-container {
        padding-top: 1.8rem;
        padding-bottom: 2rem;
    }
    /* KPI 카드 스타일 */
    .metric-card {
        background: #ffffff;
        border-radius: 12px;
        padding: 20px 24px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
        border: 1px solid #eef2f6;
        transition: transform 0.2s ease-in-out;
    }
    .metric-card:hover {
        transform: translateY(-2px);
    }
    .metric-title {
        font-size: 14px;
        color: #64748b;
        font-weight: 600;
        margin-bottom: 6px;
    }
    .metric-value {
        font-size: 26px;
        font-weight: 800;
        color: #0f172a;
    }
    .metric-delta-pos {
        font-size: 13px;
        font-weight: 700;
        color: #10b981;
    }
    .metric-delta-neg {
        font-size: 13px;
        font-weight: 700;
        color: #ef4444;
    }
</style>
""", unsafe_allow_html=True)

# 3. 데이터 로딩 설정
DEFAULT_FILE = "복사본 performance_260825.xlsx"
# 대체 파일명 지원 (플레이어 성능 등 번역 파일명 대비)
if not os.path.exists(DEFAULT_FILE):
    for f in os.listdir("."):
        if f.endswith(".xlsx") and ("performance" in f or "성능" in f):
            DEFAULT_FILE = f
            break

# 사이드바 설정
st.sidebar.markdown("### 🏢 FITI 상해지사")
st.sidebar.caption("글로벌 시험 & 제품평가 사업 실적")
st.sidebar.markdown("---")

uploaded_file = st.sidebar.file_uploader(
    "📂 최신 실적 엑셀 업로드", 
    type=["xlsx", "xls"],
    help="새로운 엑셀 파일을 올리면 화면이 실시간으로 갱신됩니다."
)

active_file = uploaded_file if uploaded_file is not None else DEFAULT_FILE

@st.cache_data
def load_all_sheets(file):
    xls = pd.ExcelFile(file)
    return xls.sheet_names

@st.cache_data
def get_summary_df(file):
    df_raw = pd.read_excel(file, sheet_name='종합', skiprows=1)
    df_raw.columns = ['구분', '세부구분', '매출_25', '매출_26', '증감수수료', '증감율']
    return df_raw.dropna(subset=['매출_25']).copy()

@st.cache_data
def get_sheet_df(file, s_name):
    return pd.read_excel(file, sheet_name=s_name)

try:
    sheet_names = load_all_sheets(active_file)
    df_summary = get_summary_df(active_file)
except Exception as e:
    st.error(f"데이터를 불러오는 중 오류가 발생했습니다: {e}")
    st.stop()

# 4. 헤더 영역
st.title("📊 FITI 상해지사 사업 실적 대시보드")
st.markdown("2025년 누적 대비 **2026년 경영 실적, 고객사별 추이 및 사업 부문별 성장률** 분석 리포트입니다.")
st.markdown("<br>", unsafe_allow_html=True)

# 5. 상단 KPI 카드 지표
total_row = df_summary[df_summary['구분'] == 'TOTAL']
if not total_row.empty:
    row = total_row.iloc[0]
    rev_25 = int(row['매출_25'])
    rev_26 = int(row['매출_26'])
    diff = int(row['증감수수료'])
    rate = float(row['증감율']) * 100

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">2025년 누적 실적</div>
            <div class="metric-value">₩ {rev_25:,.0f}</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">2026년 누적 실적</div>
            <div class="metric-value">₩ {rev_26:,.0f}</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        delta_class = "metric-delta-pos" if diff >= 0 else "metric-delta-neg"
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">전년 대비 증감액</div>
            <div class="metric-value">₩ {diff:+,.0f}</div>
            <span class="{delta_class}">{'▲' if diff >= 0 else '▼'} {abs(diff):,.0f} 원</span>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        delta_class = "metric-delta-pos" if rate >= 0 else "metric-delta-neg"
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">전년 대비 성장률</div>
            <div class="metric-value">{rate:+.2f}%</div>
            <span class="{delta_class}">{'▲' if rate >= 0 else '▼'} {abs(rate):.2f}%</span>
        </div>
        """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# 6. 인터랙티브 탭 레이아웃
tab1, tab2, tab3 = st.tabs(["📈 종합 부문 분석", "🏆 부문별 Top 고객사 / 벤더", "📋 시트별 원본 데이터"])

with tab1:
    main_dept = df_summary[~df_summary['구분'].isin(['SUB TOTAL', 'TOTAL'])].copy()
    main_dept['부문'] = (main_dept['구분'].fillna('') + ' ' + main_dept['세부구분'].fillna('')).str.strip()

    c1, c2 = st.columns([6, 4])
    
    with c1:
        st.subheader("📊 부문별 25년 vs 26년 매출 비교")
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(
            name='2025년', 
            x=main_dept['부문'], 
            y=main_dept['매출_25'],
            marker_color='#94a3b8',
            text=main_dept['매출_25'],
            texttemplate='₩ %{text:,.0f}',
            textposition='outside'
        ))
        fig_bar.add_trace(go.Bar(
            name='2026년', 
            x=main_dept['부문'], 
            y=main_dept['매출_26'],
            marker_color='#2563eb',
            text=main_dept['매출_26'],
            texttemplate='₩ %{text:,.0f}',
            textposition='outside'
        ))
        fig_bar.update_layout(
            template='plotly_white',
            barmode='group',
            height=460,
            margin=dict(l=10, r=10, t=30, b=10),
            legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1)
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    with c2:
        st.subheader("🥧 2026년 부문별 매출 비중")
        fig_pie = px.pie(
            main_dept, 
            values='매출_26', 
            names='부문', 
            hole=0.55,
            color_discrete_sequence=['#1e40af', '#3b82f6', '#60a5fa', '#93c5fd', '#38bdf8', '#0ea5e9']
        )
        fig_pie.update_traces(textposition='inside', textinfo='percent+label', marker=dict(line=dict(color='#ffffff', width=2)))
        fig_pie.update_layout(template='plotly_white', height=460, margin=dict(l=10, r=10, t=30, b=10), showlegend=False)
        st.plotly_chart(fig_pie, use_container_width=True)

with tab2:
    st.subheader("🔎 사업 부문별 실적 순위")
    target_options = ['global part1 ', 'global part2', 'KC part', 'GB part', 'inspection (원단)', 'inspection (가먼트)']
    valid_options = [s for s in target_options if s in sheet_names]
    selected_sheet = st.selectbox("분석 대상 선택", valid_options if valid_options else sheet_names)

    df_detail = get_sheet_df(active_file, selected_sheet)

    if '26년 합계' in df_detail.columns and '업체명' in df_detail.columns:
        slider_col, _ = st.columns([3, 7])
        with slider_col:
            top_limit = st.slider("조회 업체 수", 5, 20, 10)

        df_rank = df_detail.dropna(subset=['26년 합계']).sort_values(by='26년 합계', ascending=False).head(top_limit)
        
        fig_hbar = px.bar(
            df_rank,
            x='26년 합계',
            y='업체명',
            orientation='h',
            text='26년 합계',
            color='증감율' if '증감율' in df_rank.columns else None,
            color_continuous_scale='Tealgrn',
            labels={'26년 합계': '2026년 매출(원)', '업체명': '업체명'},
            title=f"[{selected_sheet}] 실적 상위 {top_limit}개 사"
        )
        fig_hbar.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
        fig_hbar.update_layout(template='plotly_white', yaxis={'categoryorder':'total ascending'}, height=520)
        st.plotly_chart(fig_hbar, use_container_width=True)
    else:
        st.dataframe(df_detail, use_container_width=True)

with tab3:
    st.subheader("📋 전체 원본 데이터 테이블")
    view_sheet = st.selectbox("조회할 시트 선택", sheet_names)
    df_raw_view = get_sheet_df(active_file, view_sheet)
    st.dataframe(df_raw_view, use_container_width=True)
