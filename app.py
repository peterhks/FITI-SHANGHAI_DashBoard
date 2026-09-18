import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from statsmodels.tsa.arima.model import ARIMA
import os

# =========================================================
# 1. 페이지 테마 및 엔터프라이즈 스타일 설정
# =========================================================
st.set_page_config(page_title="시험연구원 사업실적 분석 시스템", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    html, body, [class*="css"] {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif;
    }
    .header-box {
        padding: 22px 26px;
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
        font-size: 24px;
        color: #0F172A;
        font-weight: 700;
        letter-spacing: -0.5px;
    }
    .metric-delta {
        font-size: 12px;
        margin-top: 4px;
    }
    .section-header {
        font-size: 17px;
        font-weight: 700;
        color: #1E293B;
        margin: 22px 0 12px 0;
    }
</style>
""", unsafe_allow_html=True)

# =========================================================
# 2. 데이터 로드 및 타입 안전 정제
# =========================================================
EXCEL_FILE = "복사본 performance_260825.xlsx"

def clean_dataframe(df):
    """문자형 숫자, 공백, 표준 명칭 자동 정리"""
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
    
    # 쉼표(,)가 섞인 문자열을 순수 숫자로 강제 변환
    for col in df.columns:
        if df[col].dtype == object:
            sample = df[col].dropna().astype(str)
            if sample.str.replace(',', '').str.replace('.', '', regex=False).str.replace('-', '').str.isdigit().mean() > 0.6:
                df[col] = df[col].astype(str).str.replace(',', '').str.strip()
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    return df

@st.cache_data
def load_data():
    df = None
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE)
        except Exception:
            pass
            
    if df is None:
        dates = pd.date_range(start="2024-01-01", periods=180, freq="D")
        clients = ["코오롱스포츠", "F&F", "무신사", "삼성물산", "FILA", "데상트", "POLO"]
        test_types = ["중국 GB 규격시험", "KC 안전인증", "해외 브랜드 바이어 매뉴얼", "공장 완제품 검사", "위생용품·용기포장"]
        np.random.seed(42)
        records = []
        for d in dates:
            for _ in range(np.random.randint(1, 4)):
                fee = float(np.random.randint(300, 2500) * 1000)
                discount = float(fee * np.random.choice([0, 0.05, 0.1, 0.15]))
                records.append({
                    "접수일자": d,
                    "고객사명": np.random.choice(clients),
                    "시험분류": np.random.choice(test_types),
                    "총시험수수료": fee,
                    "감면금액": discount,
                    "실청구금액": fee - discount,
                    "성적서발급건수": int(np.random.randint(1, 6))
                })
        df = pd.DataFrame(records)
        
    return clean_dataframe(df)

df = load_data()

# =========================================================
# 3. 사이드바 제어 패널
# =========================================================
st.sidebar.markdown("### 📋 분석 파라미터")
uploaded_file = st.sidebar.file_uploader("추가 데이터 업로드 (xlsx/csv)", type=["xlsx", "csv"])
if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = clean_dataframe(pd.read_csv(uploaded_file))
    else:
        df = clean_dataframe(pd.read_excel(uploaded_file))

# 컬럼 식별
date_cols = [c for c in df.columns if any(k in c.lower() for k in ["일자", "일시", "date", "년도", "년월", "년"])]
numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
categorical_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()

if not numeric_cols:
    st.error("분석 가능한 숫자 데이터 컬럼이 없습니다.")
    st.stop()

selected_date_col = st.sidebar.selectbox("기준 일자(시계열)", date_cols if date_cols else df.columns, index=0)
selected_cat_col = st.sidebar.selectbox("분석 차원(카테고리)", categorical_cols if categorical_cols else df.columns, index=0)
primary_metric = st.sidebar.selectbox("주요 분석 지표", numeric_cols, index=0)
secondary_metric = st.sidebar.selectbox("비교/보조 지표", numeric_cols, index=1 if len(numeric_cols) > 1 else 0)

# 시계열 전처리
df[selected_date_col] = pd.to_datetime(df[selected_date_col], errors='coerce')
df = df.dropna(subset=[selected_date_col]).sort_values(by=selected_date_col)

# 안전한 숫자형 캐스팅 보장
df[primary_metric] = pd.to_numeric(df[primary_metric], errors='coerce').fillna(0)
df[secondary_metric] = pd.to_numeric(df[secondary_metric], errors='coerce').fillna(0)

# =========================================================
# 4. 헤더 및 핵심 성과 지표(KPI)
# =========================================================
st.markdown("""
<div class="header-box">
    <div class="header-title">📊 글로벌 시험인증 및 품질평가 사업실적 분석 시스템</div>
    <div class="header-subtitle">시험 성적서 접수 현황 · 시계열 ARIMA 모델링 · 고객사 기여도 정밀 진단</div>
</div>
""", unsafe_allow_html=True)

kpi1, kpi2, kpi3, kpi4 = st.columns(4)
total_val = float(df[primary_metric].sum())
mean_val = float(df[primary_metric].mean())
val_std = float(df[primary_metric].std()) if len(df) > 1 else 0.0

start_val = df[primary_metric].iloc[0] if len(df) > 0 else 0
end_val = df[primary_metric].iloc[-1] if len(df) > 0 else 0
recent_trend = ((end_val - start_val) / (start_val + 1e-5)) * 100

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
        <div class="metric-delta" style="color: #059669;">건당/일당 평균치</div>
    </div>
    """, unsafe_allow_html=True)

with kpi3:
    st.markdown(f"""
    <div class="metric-card" style="border-left-color: #D97706;">
        <div class="metric-label">실적 변동폭 (표준편차)</div>
        <div class="metric-value">{val_std:,.0f}</div>
        <div class="metric-delta" style="color: #D97706;">데이터 분산도</div>
    </div>
    """, unsafe_allow_html=True)

with kpi4:
    color = "#DC2626" if recent_trend < 0 else "#2563EB"
    sign = "+" if recent_trend >= 0 else ""
    st.markdown(f"""
    <div class="metric-card" style="border-left-color: {color};">
        <div class="metric-label">기간 전체 성장률</div>
        <div class="metric-value">{sign}{recent_trend:.1f}%</div>
        <div class="metric-delta" style="color: {color};">시계열 증감세</div>
    </div>
    """, unsafe_allow_html=True)

st.write("")

# =========================================================
# 5. 시계열 분석 & ARIMA 통계 모델링
# =========================================================
st.markdown('<div class="section-header">📈 시계열 실적 추이 및 ARIMA 모델 기반 예측 분석</div>', unsafe_allow_html=True)

with st.expander("⚙️ 시계열 파라미터 제어 (ARIMA p, d, q)", expanded=False):
    p_col, d_col, q_col, step_col = st.columns(4)
    p = p_col.number_input("자기회귀 차수 (p)", min_value=0, max_value=5, value=1)
    d = d_col.number_input("차분 차수 (d)", min_value=0, max_value=2, value=1)
    q = q_col.number_input("이동평균 차수 (q)", min_value=0, max_value=5, value=0)
    forecast_days = step_col.slider("향후 예측 일수 (Days)", 7, 90, 30)

ts_series = df.set_index(selected_date_col).resample('D')[primary_metric].sum().bfill().ffill()

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
    st.info(f"시계열 분석 안내: {e}")

# =========================================================
# 6. 다차원 시각화 (이중 축 복합 차트 + 기여도 펀넬)
# =========================================================
col_viz1, col_viz2 = st.columns([6, 4])

with col_viz1:
    st.markdown(f'<div class="section-header">📊 주간 복합 추세 ({primary_metric} vs {secondary_metric})</div>', unsafe_allow_html=True)
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
# 7. 지능형 자동 진단 리포트
# =========================================================
st.markdown('<div class="section-header">📑 지능형 자동 통계 및 품질 진단 리포트</div>', unsafe_allow_html=True)
rep_c1, rep_c2 = st.columns(2)

with rep_c1:
    st.markdown("**1) 주요 측정 지표 통계 요약표**")
    stats_table = df[[primary_metric, secondary_metric]].describe().T[['count', 'mean', 'std', 'min', '50%', 'max']]
    stats_table.columns = ['데이터 건수', '평균값', '표준편차', '최솟값', '중앙값', '최댓값']
    st.dataframe(stats_table.style.format("{:,.1f}"), use_container_width=True)

with rep_c2:
    st.markdown("**2) 실적 이상 패턴 및 상관도 진단**")
    u_bound = mean_val + 2 * val_std
    l_bound = max(0, mean_val - 2 * val_std)
    anomalies = df[(df[primary_metric] > u_bound) | (df[primary_metric] < l_bound)]
    
    top_share_cat = cat_agg.iloc[0][selected_cat_col] if len(cat_agg) > 0 else "-"
    top_share_ratio = (cat_agg.iloc[0][primary_metric] / (total_val + 1e-5)) * 100 if len(cat_agg) > 0 else 0
    corr_val = df[primary_metric].corr(df[secondary_metric]) if len(df) > 1 else 0
    
    st.success(f"📌 **핵심 점유 부문**: **`{top_share_cat}`**이(가) 전체 실적의 **{top_share_ratio:.1f}%**를 차지하고 있습니다.")
    if len(anomalies) > 0:
        st.warning(f"⚠️ **변동성 모니터링**: 2시그마(신뢰수준 95%) 임계값을 벗어난 특이 발생 건이 총 **{len(anomalies)}건** 감지되었습니다.")
    else:
        st.info("✅ **안정성 확인**: 데이터가 통계적 관리한계선(Control Limit) 내에서 안정적으로 유지되고 있습니다.")
    st.info(f"🔗 **상관 지표 평가**: `{primary_metric}`와 `{secondary_metric}` 간의 상관계수는 **{corr_val:.2f}**입니다.")
