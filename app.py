import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from statsmodels.tsa.arima.model import ARIMA
import os

# =========================================================
# 1. 페이지 설정 및 상단 헤더 커스텀 CSS
# =========================================================
st.set_page_config(
    page_title="FITI 시험연구원 - 실적 및 품질 분석 시스템",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    html, body, [class*="css"] {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif;
    }
    
    /* 상단 네이비 공식 배너 스타일 */
    .fiti-header {
        background-color: #003876;
        padding: 22px 30px;
        border-radius: 6px;
        display: flex;
        align-items: center;
        gap: 28px;
        margin-bottom: 16px;
        color: #FFFFFF;
        box-shadow: 0 3px 10px rgba(0, 56, 118, 0.15);
    }
    .fiti-logo-area {
        display: flex;
        align-items: center;
        gap: 12px;
        border-right: 1.5px solid rgba(255, 255, 255, 0.25);
        padding-right: 28px;
    }
    .fiti-logo-text {
        font-size: 26px;
        font-weight: 900;
        letter-spacing: -0.5px;
        color: #FFFFFF;
    }
    .fiti-sub-logo {
        display: flex;
        flex-direction: column;
        line-height: 1.2;
    }
    .fiti-sub-logo-kr {
        font-size: 13px;
        font-weight: 700;
        color: #FFFFFF;
    }
    .fiti-sub-logo-en {
        font-size: 10px;
        color: #B0C4DE;
    }
    .fiti-title-area {
        display: flex;
        flex-direction: column;
        line-height: 1.35;
    }
    .fiti-main-title {
        font-size: 20px;
        font-weight: 800;
        color: #FFFFFF;
        margin-bottom: 3px;
    }
    .fiti-en-title {
        font-size: 12px;
        color: #D0E1FD;
        font-weight: 400;
    }
    .fiti-dept-title {
        font-size: 11px;
        color: #A3C7F7;
        margin-top: 2px;
    }

    /* 카드 및 메트릭 커스텀 */
    .stMetric {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 12px 16px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
    }
</style>
""", unsafe_allow_html=True)

# =========================================================
# 2. FITI 공식 상단 배너 헤더
# =========================================================
st.markdown("""
<div class="fiti-header">
    <div class="fiti-logo-area">
        <div class="fiti-logo-text">FITI</div>
        <div class="fiti-sub-logo">
            <span class="fiti-sub-logo-kr">FITI 시험연구원</span>
            <span class="fiti-sub-logo-en">FITI Testing & Research Institute</span>
        </div>
    </div>
    <div class="fiti-title-area">
        <div class="fiti-main-title">사업실적 및 품질평가 종합 분석 시스템</div>
        <div class="fiti-en-title">Business Performance & Quality Analytics System</div>
        <div class="fiti-dept-title">사업팀 / 제품평가팀 (Products Evaluation Team)</div>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================
# 3. 사용 안내 (Expander 아코디언 카드)
# =========================================================
with st.expander("ℹ️ 사용 안내", expanded=True):
    st.markdown("""
    * **데이터 자동 정제**: 엑셀 또는 CSV 파일을 업로드하면 금액 내 콤마(`,`) 제거 및 주요 항목명을 표준 규격으로 자동 변환합니다.
    * **시계열 추이 & ARIMA 예측**: 기간별 실적 추세를 파악하고, 통계 모델($p, d, q$) 파라미터를 조절하여 향후 수요 및 실적을 예측합니다.
    * **다차원 인터랙티브 시각화**: 이중 축 복합 차트와 기여도 순위 펀넬(Funnel)을 통해 항목별/고객사별 핵심 성과를 확인합니다.
    * **이상 징후 자동 진단**: 관리한계선(±2σ) 기반 통계 이상치와 최대 기여 부문을 자동 산출하여 보고서 형태로 제공합니다.
    """)

# =========================================================
# 4. 데이터 로드 및 전처리
# =========================================================
EXCEL_FILE = "복사본 performance_260825.xlsx"

@st.cache_data
def load_sample_data():
    dates = pd.date_range(start="2024-01-01", periods=120, freq="D")
    categories = ["중국 GB 규격시험", "KC 안전인증", "바이어 매뉴얼", "공장 완제품검사", "위생용품·기구용기"]
    np.random.seed(42)
    
    data = []
    base_val = 50000
    for d in dates:
        base_val += np.random.randint(-1500, 2000)
        cat = np.random.choice(categories)
        cost = base_val * np.random.uniform(0.65, 0.8)
        data.append({
            "접수일자": d,
            "시험분류": cat,
            "수수료금액": float(max(10000, base_val)),
            "감면금액": float(max(5000, cost)),
            "성적서건수": int(np.random.randint(10, 100))
        })
    return pd.DataFrame(data)

# 사이드바 데이터 컨트롤
st.sidebar.title("데이터 설정 & 필터")
uploaded_file = st.sidebar.file_uploader("로컬 데이터 파일 업로드 (CSV/Excel)", type=["csv", "xlsx"])

if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
elif os.path.exists(EXCEL_FILE):
    df = pd.read_excel(EXCEL_FILE)
else:
    df = load_sample_data()

# 쉼표(,) 포함 숫자 자동 변환 (TypeError 완벽 방지)
for col in df.columns:
    if df[col].dtype == object:
        sample = df[col].dropna().astype(str)
        if sample.str.replace(',', '').str.replace('.', '', regex=False).str.replace('-', '').str.isdigit().mean() > 0.5:
            df[col] = df[col].astype(str).str.replace(',', '').str.strip()
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

# =========================================================
# 5. 사이드바 변수 매핑
# =========================================================
st.sidebar.markdown("---")
st.sidebar.subheader("분석 차원 매핑")

num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
cat_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()

date_col = st.sidebar.selectbox("시계열 기준 컬럼", df.columns, index=0)
cat_col = st.sidebar.selectbox("카테고리/차원 컬럼", cat_cols if cat_cols else df.columns, index=0)
val_col = st.sidebar.selectbox("핵심 측정값 (주요 지표)", num_cols, index=0)
sec_val_col = st.sidebar.selectbox("보조 측정값 (비교 지표)", num_cols, index=1 if len(num_cols) > 1 else 0)

# 시계열 타입 보정
df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
df = df.dropna(subset=[date_col]).sort_values(by=date_col)
df[val_col] = pd.to_numeric(df[val_col], errors='coerce').fillna(0)
df[sec_val_col] = pd.to_numeric(df[sec_val_col], errors='coerce').fillna(0)

# =========================================================
# 6. 상단 핵심 성과 지표 (KPI Cards)
# =========================================================
kpi1, kpi2, kpi3, kpi4 = st.columns(4)
total_val = float(df[val_col].sum())
mean_val = float(df[val_col].mean())
val_std = float(df[val_col].std()) if len(df) > 1 else 0.0

start_v = df[val_col].iloc[0] if len(df) > 0 else 0
end_v = df[val_col].iloc[-1] if len(df) > 0 else 0
recent_trend = ((end_v - start_v) / (start_v + 1e-5)) * 100

with kpi1:
    st.metric(label=f"총 {val_col}", value=f"{total_val:,.0f}")
with kpi2:
    st.metric(label=f"일평균 {val_col}", value=f"{mean_val:,.1f}")
with kpi3:
    st.metric(label="변동성 (표준편차)", value=f"{val_std:,.1f}")
with kpi4:
    st.metric(label="기간 전체 증감률", value=f"{recent_trend:+.2f}%")

st.markdown("---")

# =========================================================
# 7. 시계열 분석 & ARIMA 통계 모델
# =========================================================
st.subheader("1. 시계열 추이 분석 및 통계 예측 (Time-Series & ARIMA Forecast)")

with st.expander("⚙️ 시계열 파라미터 설정 (ARIMA)", expanded=False):
    c1, c2, c3, c4 = st.columns(4)
    p = c1.number_input("자기회귀 차수 (p)", min_value=0, max_value=5, value=1)
    d = c2.number_input("차분 차수 (d)", min_value=0, max_value=2, value=1)
    q = c3.number_input("이동평균 차수 (q)", min_value=0, max_value=5, value=0)
    forecast_steps = c4.slider("향후 예측 일수", 7, 60, 30)

ts_df = df.set_index(date_col).resample('D')[val_col].mean().bfill().ffill()

try:
    model = ARIMA(ts_df, order=(p, d, q))
    model_fit = model.fit()
    forecast = model_fit.forecast(steps=forecast_steps)
    forecast_dates = pd.date_range(start=ts_df.index[-1] + pd.Timedelta(days=1), periods=forecast_steps, freq='D')
    
    fig_ts = go.Figure()
    fig_ts.add_trace(go.Scatter(x=ts_df.index, y=ts_df.values, mode='lines', name=f'실측 {val_col}', line=dict(color='#003876', width=2.2)))
    fig_ts.add_trace(go.Scatter(x=forecast_dates, y=forecast, mode='lines+markers', name='ARIMA 예측선', line=dict(color='#E63946', dash='dash')))
    fig_ts.update_layout(height=400, margin=dict(l=20, r=20, t=30, b=20), hovermode='x unified', template='plotly_white')
    st.plotly_chart(fig_ts, use_container_width=True)
except Exception as e:
    st.info(f"시계열 분석 안내: {e}")

# =========================================================
# 8. 다차원 시각화 (이중 축 복합 차트 + 기여도 Funnel)
# =========================================================
st.subheader("2. 다차원 시각화 (Multidimensional Visualizations)")
col_a, col_b = st.columns([6, 4])

with col_a:
    agg_df = df.groupby(pd.Grouper(key=date_col, freq='W'))[[val_col, sec_val_col]].sum().reset_index()
    fig_combo = go.Figure()
    fig_combo.add_trace(go.Bar(x=agg_df[date_col], y=agg_df[val_col], name=val_col, marker_color='#2A62D6', opacity=0.85))
    fig_combo.add_trace(go.Scatter(x=agg_df[date_col], y=agg_df[sec_val_col], name=sec_val_col, yaxis='y2', line=dict(color='#F5A623', width=3)))
    fig_combo.update_layout(
        title="주간 복합 추세 분석 (Dual-Axis)",
        yaxis=dict(title=val_col),
        yaxis2=dict(title=sec_val_col, overlaying='y', side='right'),
        legend=dict(x=0, y=1.12, orientation='h'),
        height=380,
        template='plotly_white'
    )
    st.plotly_chart(fig_combo, use_container_width=True)

with col_b:
    funnel_df = df.groupby(cat_col)[val_col].sum().reset_index().sort_values(by=val_col, ascending=False)
    fig_funnel = go.Figure(go.Funnel(
        y=funnel_df[cat_col],
        x=funnel_df[val_col],
        textinfo="value+percent initial",
        marker={"color": ["#003876", "#1D5BB0", "#458BE8", "#82B7FC", "#BDD9FE"][:len(funnel_df)]}
    ))
    fig_funnel.update_layout(title="카테고리별 기여도 (Funnel 차트)", height=380, margin=dict(t=40, b=20, l=10, r=10))
    st.plotly_chart(fig_funnel, use_container_width=True)

# =========================================================
# 9. 지능형 통계 및 리포트 자동 진단
# =========================================================
st.subheader("3. 지능형 자동 진단 리포트 (Automated Analytics)")
col_rep1, col_rep2 = st.columns(2)

with col_rep1:
    st.markdown("**통계 지표 요약**")
    summary_stats = df[[val_col, sec_val_col]].describe().T[['mean', 'std', 'min', '50%', 'max']]
    summary_stats.columns = ['평균', '표준편차', '최소값', '중앙값', '최대값']
    st.dataframe(summary_stats.style.format("{:,.2f}"), use_container_width=True)

with col_rep2:
    st.markdown("**이상치 및 패턴 자동 진단**")
    upper_bound = mean_val + 2 * val_std
    lower_bound = max(0, mean_val - 2 * val_std)
    outliers = df[(df[val_col] > upper_bound) | (df[val_col] < lower_bound)]
    
    top_cat = funnel_df.iloc[0][cat_col] if len(funnel_df) > 0 else "-"
    top_cat_share = (funnel_df.iloc[0][val_col] / (total_val + 1e-5)) * 100 if len(funnel_df) > 0 else 0
    
    st.success(f"**핵심 점유군**: `{top_cat}`이 전체 실적의 **{top_cat_share:.1f}%**를 견인하고 있습니다.")
    if len(outliers) > 0:
        st.warning(f"**변동성 감지**: 정상 범주(±2σ)를 벗어난 이상 패턴이 **{len(outliers)}건** 감지되었습니다.")
    else:
        st.info("**변동성 감지**: 데이터가 통계적 관리한계 내에서 안정적으로 제어되고 있습니다.")
    
    corr = df[val_col].corr(df[sec_val_col]) if len(df) > 1 else 0
    st.write(f"- `{val_col}`과 `{sec_val_col}`의 상관계수: **{corr:.2f}**")
