import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from statsmodels.tsa.arima.model import ARIMA
import os

# =========================================================
# 1. 페이지 테마 및 세련된 비즈니스 스타일 CSS
# =========================================================
st.set_page_config(page_title="시험연구원 사업실적 분석 시스템", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    /* 폰트 및 배경 기본 스타일링 */
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    html, body, [class*="css"] {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif;
    }
    
    /* 상단 헤더 컨테이너 */
    .header-box {
        padding: 20px 24px;
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
        border-radius: 12px;
        color: #FFFFFF;
        margin-bottom: 24px;
        box-shadow: 0 4px 12px rgba(30, 58, 138, 0.15);
    }
    .header-title {
        font-size: 24px;
        font-weight: 700;
        margin: 0;
        color: #FFFFFF;
    }
    .header-subtitle {
        font-size: 14px;
        margin-top: 6px;
        color: #DBEAFE;
    }

    /* KPI 카드 스타일 */
    .metric-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 18px 20px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        border-left: 5px solid #2563EB;
    }
    .metric-label {
        font-size: 13px;
        color: #64748B;
        font-weight: 600;
        margin-bottom: 6px;
    }
    .metric-value {
        font-size: 26px;
        color: #0F172A;
        font-weight: 700;
        letter-spacing: -0.5px;
    }
    .metric-delta {
        font-size: 12px;
        margin-top: 4px;
    }

    /* 모듈 컨테이너 */
    .section-header {
        font-size: 17px;
        font-weight: 700;
        color: #1E293B;
        margin: 20px 0 12px 0;
        display: flex;
        align-items: center;
        gap: 8px;
    }
</style>
""", unsafe_allow_html=True)

# =========================================================
# 2. 데이터 로드 및 표준 명칭(용어 매핑) 전처리
# =========================================================
EXCEL_FILE = "복사본 performance_260825.xlsx"

@st.cache_data
def load_and_standardize_data():
    df = None
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE)
        except Exception:
            pass
            
    # 파일이 없거나 로드 실패 시 시험기관 표준 데이터셋 생성
    if df is None:
        dates = pd.date_range(start="2024-01-01", periods=180, freq="D")
        clients = ["코오롱스포츠", "F&F", "무신사", "삼성물산", "FILA", "데상트", "POLO"]
        test_types = ["중국 GB 규격시험", "KC 안전인증", "해외 브랜드 바이어 매뉴얼", "공장 완제품 검사", "위생용품·용기포장"]
        np.random.seed(42)
        
        records = []
        for d in dates:
            for _ in range(np.random.randint(1, 4)):
                fee = np.random.randint(300, 2500) * 1000
                discount = fee * np.random.choice([0, 0.05, 0.1, 0.15])
                records.append({
                    "접수일자": d,
                    "주요고객사": np.random.choice(clients),
                    "시험분류": np.random.choice(test_types),
                    "총시험수수료": fee,
                    "감면금액": discount,
                    "실청구금액": fee - discount,
                    "성적서발급건수": np.random.randint(1, 6)
                })
        df = pd.DataFrame(records)
        
    # 기존 파일의 열 명칭을 표준 비즈니스 용어로 자동 보정 매핑
    rename_rules = {
        "수입금액": "총시험수수료",
        "공급액": "총시험수수료",
        "자산금액": "감면금액",
        "감면액": "감면금액",
        "년도": "접수연도",
        "년월": "접수연월",
        "일자": "접수일자",
        "일시": "접수일자",
        "항목": "시험분류",
        "상품명": "시험항목",
        "분류": "업무구분",
        "지점": "사업팀/센터",
        "조합명": "고객사명",
        "대분류": "시험분류"
    }
    df = df.rename(columns={c: rename_rules[c] for c in df.columns if c in rename_rules})
    return df

raw_df = load_and_standardize_data()

# =========================================================
# 3. 사이드바 인터랙티브 필터 패널
# =========================================================
st.sidebar.markdown("### 📋 분석 조건 설정")

# 파일 업로더 지원
uploaded_file = st.sidebar.file_uploader("추가 실적 파일 업로드 (xlsx/csv)", type=["xlsx", "csv"])
if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        raw_df = pd.read_csv(uploaded_file)
    else:
        raw_df = pd.read_excel(uploaded_file)

# 컬럼 식별
date_candidates = [c for c in raw_df.columns if "일" in c or "date" in c.lower() or "년" in c]
num_cols = raw_df.select_dtypes(include=[np.number]).columns.tolist()
cat_cols = raw_df.select_dtypes(include=['object', 'category']).columns.tolist()

selected_date_col = st.sidebar.selectbox("기준 일자(시계열)", date_candidates, index=0 if date_candidates else 0)
selected_cat_col = st.sidebar.selectbox("핵심 분석 분류(차원)", cat_cols, index=0 if cat_cols else 0)
primary_metric = st.sidebar.selectbox("주요 분석 지표", num_cols, index=0 if num_cols else 0)
secondary_metric = st.sidebar.selectbox("비교/보조 지표", num_cols, index=1 if len(num_cols) > 1 else 0)

# 시계열 변환
raw_df[selected_date_col] = pd.to_datetime(raw_df[selected_date_col], errors='coerce')
df = raw_df.dropna(subset=[selected_date_col]).sort_values(by=selected_date_col)

# =========================================================
# 4. 상단 헤더 및 핵심 성과 지표(KPI)
# =========================================================
st.markdown("""
<div class="header-box">
    <div class="header-title">📊 글로벌 시험인증 및 품질평가 사업실적 분석 시스템</div>
    <div class="header-subtitle">실시간 시험 성적서 접수 현황 · 시계열 예측(ARIMA) · 고객사별 기여도 통합 대시보드</div>
</div>
""", unsafe_allow_html=True)

kpi1, kpi2, kpi3, kpi4 = st.columns(4)
total_val = df[primary_metric].sum()
mean_val = df[primary_metric].mean()
val_std = df[primary_metric].std()
recent_trend = ((df[primary_metric].iloc[-1] - df[primary_metric].iloc[0]) / (df[primary_metric].iloc[0] + 1e-5)) * 100

with kpi1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">총 {primary_metric} 합계</div>
        <div class="metric-value">{total_val:,.0f} 원</div>
        <div class="metric-delta" style="color: #2563EB;">누적 실적 집계</div>
    </div>
    """, unsafe_allow_html=True)

with kpi2:
    st.markdown(f"""
    <div class="metric-card" style="border-left-color: #059669;">
        <div class="metric-label">평균 {primary_metric}</div>
        <div class="metric-value">{mean_val:,.0f} 원</div>
        <div class="metric-delta" style="color: #059669;">건당/일당 평균 산출치</div>
    </div>
    """, unsafe_allow_html=True)

with kpi3:
    st.markdown(f"""
    <div class="metric-card" style="border-left-color: #D97706;">
        <div class="metric-label">실적 변동폭 (표준편차)</div>
        <div class="metric-value">{val_std:,.0f}</div>
        <div class="metric-delta" style="color: #D97706;">신뢰구간 정밀도 지표</div>
    </div>
    """, unsafe_allow_html=True)

with kpi4:
    color = "#DC2626" if recent_trend < 0 else "#2563EB"
    sign = "+" if recent_trend >= 0 else ""
    st.markdown(f"""
    <div class="metric-card" style="border-left-color: {color};">
        <div class="metric-label">전체 구간 성장률</div>
        <div class="metric-value">{sign}{recent_trend:.1f}%</div>
        <div class="metric-delta" style="color: {color};">추세 방향성 지표</div>
    </div>
    """, unsafe_allow_html=True)

st.write("")

# =========================================================
# 5. 시계열 분석 & ARIMA 통계 예측 패널
# =========================================================
st.markdown('<div class="section-header">📈 시계열 실적 추이 및 ARIMA 모델 기반 예측 분석</div>', unsafe_allow_html=True)

with st.expander("⚙️ 시계열 모델 세부 파라미터 제어 (ARIMA p, d, q)", expanded=False):
    p_col, d_col, q_col, step_col = st.columns(4)
    p = p_col.number_input("자기회귀 차수 (p)", min_value=0, max_value=5, value=1)
    d = d_col.number_input("차분 차수 (d)", min_value=0, max_value=2, value=1)
    q = q_col.number_input("이동평균 차수 (q)", min_value=0, max_value=5, value=0)
    forecast_days = step_col.slider("향후 예측 일수 (Days)", 7, 90, 30)

# 시계열 리샘플링
ts_series = df.set_index(selected_date_col).resample('D')[primary_metric].sum().fillna(method='ffill')

try:
    model = ARIMA(ts_series, order=(p, d, q))
    model_res = model.fit()
    forecast = model_res.forecast(steps=forecast_days)
    forecast_idx = pd.date_range(start=ts_series.index[-1] + pd.Timedelta(days=1), periods=forecast_days, freq='D')
    
    fig_ts = go.Figure()
    fig_ts.add_trace(go.Scatter(
        x=ts_series.index, y=ts_series.values,
        mode='lines', name=f'실측 {primary_metric}',
        line=dict(color='#2563EB', width=2.5)
    ))
    fig_ts.add_trace(go.Scatter(
        x=forecast_idx, y=forecast,
        mode='lines+markers', name='ARIMA 통계 예측선',
        line=dict(color='#DC2626', width=2.5, dash='dash')
    ))
    fig_ts.update_layout(
        template='plotly_white',
        height=380,
        margin=dict(l=20, r=20, t=30, b=20),
        hovermode='x unified',
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    st.plotly_chart(fig_ts, use_container_width=True)
except Exception as e:
    st.warning(f"시계열 모델 산출 정보: {e}")

# =========================================================
# 6. 다차원 정밀 시각화 (이중 축 복합 차트 + 기여도 펀넬)
# =========================================================
col_viz1, col_viz2 = st.columns([6, 4])

with col_viz1:
    st.markdown(f'<div class="section-header">📊 주간 복합 실적 ({primary_metric} vs {secondary_metric})</div>', unsafe_allow_html=True)
    weekly_agg = df.groupby(pd.Grouper(key=selected_date_col, freq='W'))[[primary_metric, secondary_metric]].sum().reset_index()
    
    fig_dual = go.Figure()
    fig_dual.add_trace(go.Bar(
        x=weekly_agg[selected_date_col], y=weekly_agg[primary_metric],
        name=primary_metric, marker_color='#3B82F6', opacity=0.85
    ))
    fig_dual.add_trace(go.Scatter(
        x=weekly_agg[selected_date_col], y=weekly_agg[secondary_metric],
        name=secondary_metric, yaxis='y2',
        line=dict(color='#F59E0B', width=3)
    ))
    fig_dual.update_layout(
        template='plotly_white',
        height=360,
        margin=dict(l=20, r=20, t=20, b=20),
        yaxis=dict(title=primary_metric),
        yaxis2=dict(title=secondary_metric, overlaying='y', side='right'),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    st.plotly_chart(fig_dual, use_container_width=True)

with col_viz2:
    st.markdown(f'<div class="section-header">🎯 {selected_cat_col}별 실적 기여도 (Ranked Funnel)</div>', unsafe_allow_html=True)
    cat_agg = df.groupby(selected_cat_col)[primary_metric].sum().reset_index().sort_values(by=primary_metric, ascending=False)
    
    fig_cat = go.Figure(go.Funnel(
        y=cat_agg[selected_cat_col],
        x=cat_agg[primary_metric],
        textinfo="value+percent initial",
        marker=dict(colors=["#1E3A8A", "#2563EB", "#3B82F6", "#60A5FA", "#93C5FD", "#BFDBFE"])
    ))
    fig_cat.update_layout(
        template='plotly_white',
        height=360,
        margin=dict(l=20, r=20, t=20, b=20)
    )
    st.plotly_chart(fig_cat, use_container_width=True)

# =========================================================
# 7. 전문가 자동 진단 리포트 (통계 진단 및 품질 관리 인사이트)
# =========================================================
st.markdown('<div class="section-header">📑 지능형 자동 품질 & 성과 진단 리포트</div>', unsafe_allow_html=True)
rep_c1, rep_c2 = st.columns(2)

with rep_c1:
    st.markdown("**1) 주요 측정 지표 통계 요약표**")
    stats_table = df[[primary_metric, secondary_metric]].describe().T[['count', 'mean', 'std', 'min', '50%', 'max']]
    stats_table.columns = ['표본수', '평균값', '표준편차', '최솟값', '중위수(50%)', '최댓값']
    st.dataframe(stats_table.style.format("{:,.1f}"), use_container_width=True)

with rep_c2:
    st.markdown("**2) 실적 이상 징후 및 점유 분석**")
    
    # 2-Sigma 이상치 탐지
    u_bound = mean_val + 2 * val_std
    l_bound = max(0, mean_val - 2 * val_std)
    anomalies = df[(df[primary_metric] > u_bound) | (df[primary_metric] < l_bound)]
    
    top_share_cat = cat_agg.iloc[0][selected_cat_col]
    top_share_ratio = (cat_agg.iloc[0][primary_metric] / total_val) * 100
    corr_val = df[primary_metric].corr(df[secondary_metric])
    
    st.success(f"📌 **최대 실적 기여군**: **`{top_share_cat}`** 부문이 전체의 **{top_share_ratio:.1f}%** 비중을 차지하고 있습니다.")
    if len(anomalies) > 0:
        st.warning(f"⚠️ **변동성 모니터링**: 2시그마(신뢰수준 95%) 임계값을 벗어난 특이 발생 건이 총 **{len(anomalies)}건** 확인되었습니다.")
    else:
        st.info("✅ **안정성 검증**: 측정치가 통계적 관리한계선(Control Limit) 내에서 균형 있게 유지되고 있습니다.")
    st.info(f"🔗 **상관 지표 평가**: `{primary_metric}`와 `{secondary_metric}` 간의 상관계수는 **{corr_val:.2f}**로 확인됩니다.")
