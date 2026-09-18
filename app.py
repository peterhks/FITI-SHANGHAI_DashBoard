import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from statsmodels.tsa.arima.model import ARIMA

# 1. 페이지 레이아웃 및 스타일 설정
st.set_page_config(page_title="AI Data Analytics Dashboard", layout="wide")

st.markdown("""
<style>
    .metric-box {
        background-color: #f8f9fa;
        padding: 16px;
        border-radius: 10px;
        border: 1px solid #e9ecef;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

# 2. 샘플 데이터셋 자동 생성 함수 (업로드 파일 없을 시 구동)
@st.cache_data
def load_sample_data():
    dates = pd.date_range(start="2024-01-01", periods=120, freq="D")
    categories = ["수산물", "가공식품", "지역특산물", "공산품"]
    np.random.seed(42)
    
    data = []
    base_val = 50000
    for d in dates:
        base_val += np.random.randint(-1500, 2000)
        cat = np.random.choice(categories)
        cost = base_val * np.random.uniform(0.65, 0.8)
        data.append({
            "일자": d,
            "대분류": cat,
            "공급액": max(10000, base_val),
            "감면액": max(5000, cost),
            "수량": np.random.randint(10, 100)
        })
    return pd.DataFrame(data)

# 사이드바: 데이터 업로드 및 조건 제어
st.sidebar.title("데이터 설정 & 필터")
uploaded_file = st.sidebar.file_uploader("로컬 데이터 파일 업로드 (CSV/Excel)", type=["csv", "xlsx"])

if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
else:
    df = load_sample_data()

# 3. 사이드바: 분석 변수 매핑
st.sidebar.markdown("---")
st.sidebar.subheader("분석 차원 매핑")
date_col = st.sidebar.selectbox("시계열 기준 컬럼", df.columns, index=0)
cat_col = st.sidebar.selectbox("카테고리/차원 컬럼", df.columns, index=1 if len(df.columns) > 1 else 0)
val_col = st.sidebar.selectbox("핵심 측정값 (주요 지표)", df.select_dtypes(include=[np.number]).columns, index=0)
sec_val_col = st.sidebar.selectbox("보조 측정값 (비교 지표)", df.select_dtypes(include=[np.number]).columns, index=1 if len(df.select_dtypes(include=[np.number]).columns) > 1 else 0)

# 시계열 전처리
df[date_col] = pd.to_datetime(df[date_col])
df = df.sort_values(by=date_col)

# 4. 상단 KPI 카드 영역
st.title("지능형 다차원 분석 대시보드")
st.caption("Auto-Analytics & Multidimensional Visualization System")

kpi1, kpi2, kpi3, kpi4 = st.columns(4)
total_val = df[val_col].sum()
mean_val = df[val_col].mean()
val_std = df[val_col].std()
recent_trend = (df[val_col].iloc[-1] - df[val_col].iloc[0]) / df[val_col].iloc[0] * 100

with kpi1:
    st.metric(label=f"총 {val_col}", value=f"{total_val:,.0f}")
with kpi2:
    st.metric(label=f"일평균 {val_col}", value=f"{mean_val:,.1f}")
with kpi3:
    st.metric(label="변동성 (표준편차)", value=f"{val_std:,.1f}")
with kpi4:
    st.metric(label="기간 전체 증감률", value=f"{recent_trend:+.2f}%")

st.markdown("---")

# 5. 시계열 분석 & 모달(팝업)형 예측 설정 영역
st.subheader("1. 시계열 고급 분석 (Time-Series & ARIMA Forecast)")

# 상단 인터랙티브 제어 패널
with st.expander("시계열 파라미터 제어 (ARIMA 설정)", expanded=False):
    c1, c2, c3, c4 = st.columns(4)
    p = c1.number_input("자기회귀 차수 (p)", min_value=0, max_value=5, value=1)
    d = c2.number_input("차분 차수 (d)", min_value=0, max_value=2, value=1)
    q = c3.number_input("이동평균 차수 (q)", min_value=0, max_value=5, value=0)
    forecast_steps = c4.slider("향후 예측 기간", 7, 60, 30)

# ARIMA 모델 실행 및 시계열 차트 생성
ts_df = df.set_index(date_col).resample('D')[val_col].mean().interpolate(method='linear')

try:
    model = ARIMA(ts_df, order=(p, d, q))
    model_fit = model.fit()
    forecast = model_fit.forecast(steps=forecast_steps)
    forecast_dates = pd.date_range(start=ts_df.index[-1] + pd.Timedelta(days=1), periods=forecast_steps, freq='D')
    
    fig_ts = go.Figure()
    fig_ts.add_trace(go.Scatter(x=ts_df.index, y=ts_df.values, mode='lines', name=f'실측 {val_col}', line=dict(color='#2A62D6', width=2)))
    fig_ts.add_trace(go.Scatter(x=forecast_dates, y=forecast, mode='lines+markers', name=f'ARIMA 예측선', line=dict(color='#E63946', dash='dash')))
    fig_ts.update_layout(height=420, margin=dict(l=20, r=20, t=30, b=20), hovermode='x unified', template='plotly_white')
    st.plotly_chart(fig_ts, use_container_width=True)
except Exception as e:
    st.error(f"ARIMA 연산 중 오류가 발생했습니다. 데이터 또는 차수를 재설정하세요: {e}")

# 6. 다차원 시각화 영역 (복합 차트, Funnel, Sunburst)
st.subheader("2. 다차원 시각화 (Multidimensional Visualizations)")
col_a, col_b = st.columns([6, 4])

with col_a:
    # 이중 축 복합 차트 (Bar + Line)
    agg_df = df.groupby(pd.Grouper(key=date_col, freq='W'))[[val_col, sec_val_col]].sum().reset_index()
    fig_combo = go.Figure()
    fig_combo.add_trace(go.Bar(x=agg_df[date_col], y=agg_df[val_col], name=val_col, marker_color='#4A90E2', opacity=0.8))
    fig_combo.add_trace(go.Scatter(x=agg_df[date_col], y=agg_df[sec_val_col], name=sec_val_col, yaxis='y2', line=dict(color='#F5A623', width=3)))
    fig_combo.update_layout(
        title="주간 복합 추세 분석 (Dual-Axis)",
        yaxis=dict(title=val_col),
        yaxis2=dict(title=sec_val_col, overlaying='y', side='right'),
        legend=dict(x=0, y=1.1, orientation='h'),
        height=380,
        template='plotly_white'
    )
    st.plotly_chart(fig_combo, use_container_width=True)

with col_b:
    # 펀넬(Funnel) 및 점유율 파이 차트
    funnel_df = df.groupby(cat_col)[val_col].sum().reset_index().sort_values(by=val_col, ascending=False)
    fig_funnel = go.Figure(go.Funnel(
        y=funnel_df[cat_col],
        x=funnel_df[val_col],
        textinfo="value+percent initial",
        marker={"color": ["#1D3557", "#457B9D", "#A8DADC", "#F1FAEE"][:len(funnel_df)]}
    ))
    fig_funnel.update_layout(title="카테고리별 기여도 (Funnel 차트)", height=380, margin=dict(t=40, b=20, l=10, r=10))
    st.plotly_chart(fig_funnel, use_container_width=True)

# 7. 자체 지능형 통계 및 리포트 자동 생성 영역
st.subheader("3. 지능형 자동 진단 리포트 (Automated Analytics)")

col_rep1, col_rep2 = st.columns(2)

with col_rep1:
    st.markdown("**통계 지표 요약**")
    summary_stats = df[[val_col, sec_val_col]].describe().T[['mean', 'std', 'min', '50%', 'max']]
    summary_stats.columns = ['평균', '표준편차', '최소값', '중앙값', '최대값']
    st.dataframe(summary_stats.style.format("{:,.2f}"), use_container_width=True)

with col_rep2:
    st.markdown("**이상치 및 패턴 자동 진단**")
    
    # 3-Sigma 기반 이상 징후 자동 추출
    upper_bound = mean_val + 2 * val_std
    lower_bound = max(0, mean_val - 2 * val_std)
    outliers = df[(df[val_col] > upper_bound) | (df[val_col] < lower_bound)]
    
    # 최고 기여 카테고리 도출
    top_cat = funnel_df.iloc[0][cat_col]
    top_cat_share = (funnel_df.iloc[0][val_col] / total_val) * 100
    
    st.success(f"**핵심 점유군**: `{top_cat}`이 전체 실적의 **{top_cat_share:.1f}%**를 견인하고 있습니다.")
    if len(outliers) > 0:
        st.warning(f"**변동성 감지**: 정상 범주(±2σ)를 벗어난 이상 패턴이 **{len(outliers)}건** 감지되었습니다.")
    else:
        st.info("**변동성 감지**: 데이터가 통계적 임계치 내에서 안정적으로 제어되고 있습니다.")
    
    # 상관계수 자동 평가
    corr = df[val_col].corr(df[sec_val_col])
    st.write(f"- `{val_col}`과 `{sec_val_col}`의 상관계수: **{corr:.2f}** ({'강한 정적 상관관계' if corr > 0.6 else '독립적 추세'})")
