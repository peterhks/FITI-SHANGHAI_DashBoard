import streamlit as st
import pandas as pd
import plotly.express as px
import os

# =========================================================
# 1. 화면 기본 설정 및 디자인 스타일
# =========================================================
st.set_page_config(
    page_title="FITI SHANGHAI 실적 분석",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    html, body, [class*="css"] {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif;
    }
    
    /* 상단 네이비 공식 배너 */
    .fiti-header {
        background-color: #003876;
        padding: 20px 26px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        gap: 20px;
        color: #FFFFFF;
        margin-bottom: 20px;
        box-shadow: 0 4px 12px rgba(0, 56, 118, 0.15);
    }
    .fiti-logo-text {
        font-size: 26px;
        font-weight: 900;
        letter-spacing: -0.5px;
        border-right: 1.5px solid rgba(255, 255, 255, 0.25);
        padding-right: 22px;
    }
    .fiti-title-main {
        font-size: 20px;
        font-weight: 800;
        margin-bottom: 3px;
    }
    .fiti-title-sub {
        font-size: 12px;
        color: #D0E1FD;
    }

    /* 입체형 KPI 카드 스타일 */
    .kpi-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 20px 22px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -2px rgba(0, 0, 0, 0.05);
        transition: transform 0.15s ease, box-shadow 0.15s ease;
        border-top: 4px solid #CBD5E1;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08);
    }
    .kpi-title {
        font-size: 13px;
        font-weight: 600;
        color: #64748B;
        margin-bottom: 8px;
    }
    .kpi-num {
        font-size: 26px;
        font-weight: 800;
        color: #0F172A;
        letter-spacing: -0.5px;
    }
    .kpi-sub {
        font-size: 12px;
        color: #94A3B8;
        margin-top: 6px;
    }
    .kpi-badge {
        display: inline-block;
        padding: 3px 8px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 700;
        margin-top: 6px;
    }
</style>
""", unsafe_allow_html=True)

# =========================================================
# 2. FITI 공식 상단 배너
# =========================================================
st.markdown("""
<div class="fiti-header">
    <div class="fiti-logo-text">FITI</div>
    <div>
        <div class="fiti-title-main">상해지사 실적 종합 분석</div>
        <div class="fiti-title-sub">상해지사 사업 실적 및 분석 시스템 | 상해지사 사업팀</div>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================
# 3. 데이터 로드 및 결측치/문자형 숫자 정제
# =========================================================
EXCEL_FILE = "복사본 performance_260825.xlsx"

@st.cache_data
def get_data():
    if os.path.exists(EXCEL_FILE):
        try:
            return pd.read_excel(EXCEL_FILE)
        except Exception:
            pass
    return pd.DataFrame({
        "바이어명": ["코오롱스포츠", "F&F", "무신사", "삼성물산", "FILA"],
        "25년 1월": [45000000, 32000000, 28000000, 19000000, 12000000],
        "26년 1월": [48000000, 31000000, 33000000, 22000000, 15000000]
    })

uploaded_file = st.sidebar.file_uploader("엑셀/CSV 파일 업로드", type=["xlsx", "csv"])
if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
else:
    df = get_data()

# 쉼표(,) 포함 문자열 숫자 자동 변환 (연산 에러 방지)
for col in df.columns:
    if df[col].dtype == object:
        cleaned = df[col].astype(str).str.replace(',', '').str.strip()
        converted = pd.to_numeric(cleaned, errors='coerce')
        if converted.notnull().mean() > 0.6:
            df[col] = converted.fillna(0)

num_cols = df.select_dtypes(include=['number']).columns.tolist()
other_cols = [c for c in df.columns if c not in num_cols]

if not num_cols:
    st.error("데이터에 분석할 수 있는 숫자(실적/금액 등) 컬럼이 없습니다.")
    st.stop()

# =========================================================
# 4. 사이드바 축 및 실적 지표 선택
# =========================================================
st.sidebar.markdown("---")
st.sidebar.subheader("📊 실적 비교 및 축 선택")

default_25_idx = next((i for i, c in enumerate(num_cols) if "25" in str(c)), 0)
default_26_idx = next((i for i, c in enumerate(num_cols) if "26" in str(c)), 1 if len(num_cols) > 1 else 0)

col_25 = st.sidebar.selectbox("2025년 실적 컬럼", num_cols, index=default_25_idx)
col_26 = st.sidebar.selectbox("2026년 실적 컬럼", num_cols, index=default_26_idx)
x_axis = st.sidebar.selectbox("기준 축 (바이어명 / 항목)", other_cols if other_cols else df.columns, index=0)

# =========================================================
# 5. 상단 실적 비교 지표 (입체 카드 디자인)
# =========================================================
total_25 = float(df[col_25].sum())
total_26 = float(df[col_26].sum())
diff_val = total_26 - total_25
diff_rate = (diff_val / total_25 * 100) if total_25 != 0 else 0.0

is_positive = diff_val >= 0
diff_color = "#E11D48" if is_positive else "#2563EB"
badge_bg = "#FFE4E6" if is_positive else "#DBEAFE"
diff_sign = "+" if is_positive else ""

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.markdown(f"""
    <div class="kpi-card" style="border-top-color: #64748B;">
        <div class="kpi-title">📅 25년 총 실적</div>
        <div class="kpi-num">{total_25:,.0f}</div>
        <div class="kpi-sub">전년 누적 집계액</div>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown(f"""
    <div class="kpi-card" style="border-top-color: #003876;">
        <div class="kpi-title">🚀 26년 총 실적</div>
        <div class="kpi-num" style="color: #003876;">{total_26:,.0f}</div>
        <div class="kpi-sub">당해 누적 집계액</div>
    </div>
    """, unsafe_allow_html=True)

with c3:
    st.markdown(f"""
    <div class="kpi-card" style="border-top-color: {diff_color};">
        <div class="kpi-title">📈 실적 증감액</div>
        <div class="kpi-num" style="color: {diff_color};">{diff_sign}{diff_val:,.0f}</div>
        <span class="kpi-badge" style="background-color: {badge_bg}; color: {diff_color};">
            전년 대비 실적차
        </span>
    </div>
    """, unsafe_allow_html=True)

with c4:
    st.markdown(f"""
    <div class="kpi-card" style="border-top-color: {diff_color};">
        <div class="kpi-title">📊 증감 퍼센트</div>
        <div class="kpi-num" style="color: {diff_color};">{diff_sign}{diff_rate:0.2f}%</div>
        <span class="kpi-badge" style="background-color: {badge_bg}; color: {diff_color};">
            전년 대비 성장률
        </span>
    </div>
    """, unsafe_allow_html=True)

st.write("")
st.markdown("---")

# =========================================================
# 6. 핵심 그래프 2종 (실적 비교 막대 & 비중 파이 차트)
# =========================================================
col_left, col_right = st.columns([6, 4])

with col_left:
    # 요청하신 타이틀로 정확히 변경
    st.subheader(f"📌 {x_axis}별 2025년 총 실적 vs 2026년 총 실적 비교")
    
    # 상위 15개 항목 집계
    chart_data = df.groupby(x_axis, as_index=False)[[col_25, col_26]].sum().sort_values(by=col_26, ascending=False).head(15)
    
    # 범례 명칭을 통일된 '2025년 총 실적', '2026년 총 실적'으로 변경
    chart_data_renamed = chart_data.rename(columns={
        col_25: "2025년 총 실적",
        col_26: "2026년 총 실적"
    })
    
    fig_bar = px.bar(
        chart_data_renamed,
        x=x_axis,
        y=["2025년 총 실적", "2026년 총 실적"],
        barmode='group',
        labels={"value": "실적금액 (원)", "variable": "실적 구분"},
        color_discrete_map={"2025년 총 실적": "#93C5FD", "2026년 총 실적": "#003876"}
    )
    
    # 막대가 뭉개지지 않도록 Y축을 0원부터 정상 표기되도록 rangemode 설정
    fig_bar.update_layout(
        height=450,
        xaxis_tickangle=-45,
        yaxis=dict(rangemode='tozero', title="실적금액 (원)"),
        template="plotly_white",
        legend=dict(orientation="h", y=1.12, x=0)
    )
    st.plotly_chart(fig_bar, use_container_width=True)

with col_right:
    st.subheader(f"🥧 2026년 {x_axis} 점유율 비중")
    fig_pie = px.pie(
        chart_data.head(8),
        names=x_axis,
        values=col_26,
        hole=0.4
    )
    fig_pie.update_traces(textposition='inside', textinfo='percent+label')
    fig_pie.update_layout(height=450, margin=dict(t=20, b=20, l=10, r=10))
    st.plotly_chart(fig_pie, use_container_width=True)

# =========================================================
# 7. 원본 데이터 확인
# =========================================================
with st.expander("📄 데이터 테이블 확인"):
    st.dataframe(df, use_container_width=True)
