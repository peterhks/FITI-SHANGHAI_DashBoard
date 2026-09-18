import streamlit as st
import pandas as pd
import plotly.express as px
import os

# 1. 화면 기본 설정
st.set_page_config(page_title="FITI SHANGHAI 실적 분석", layout="wide")

# 2. FITI 공식 상단 배너
st.markdown("""
<div style="background-color:#003876; padding:20px 24px; border-radius:8px; display:flex; align-items:center; gap:20px; color:#ffffff; margin-bottom:20px;">
    <div style="font-size:26px; font-weight:900; border-right:1px solid rgba(255,255,255,0.3); padding-right:20px;">FITI</div>
    <div>
        <div style="font-size:18px; font-weight:700;">상해지사 실적 종합 분석</div>
        <div style="font-size:12px; color:#D0E1FD;">상해지사 사업 실적 및 분석 시스템 | 상해지사 사업팀</div>
    </div>
</div>
""", unsafe_allow_html=True)

# 3. 데이터 로드 (저장소 파일 또는 샘플 데이터)
EXCEL_FILE = "복사본 performance_260825.xlsx"

@st.cache_data
def get_data():
    if os.path.exists(EXCEL_FILE):
        try:
            return pd.read_excel(EXCEL_FILE)
        except Exception:
            pass
    return pd.DataFrame({
        "구분": ["중국 GB시험", "KC인증", "바이어 매뉴얼", "공장 완제품검사", "위생용품"],
        "25년 실적": [45000000, 32000000, 28000000, 19000000, 12000000],
        "26년 실적": [48000000, 31000000, 33000000, 22000000, 15000000]
    })

uploaded_file = st.sidebar.file_uploader("엑셀/CSV 파일 업로드", type=["xlsx", "csv"])
if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
else:
    df = get_data()

# 4. 데이터 자동 정리 (문자형 숫자, 쉼표 자동 변환)
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

# 5. 간편 그래프 및 실적 비교 축 선택
st.sidebar.markdown("---")
st.sidebar.subheader("📊 실적 비교 및 축 선택")

default_25_idx = next((i for i, c in enumerate(num_cols) if "25" in str(c)), 0)
default_26_idx = next((i for i, c in enumerate(num_cols) if "26" in str(c)), 1 if len(num_cols) > 1 else 0)

col_25 = st.sidebar.selectbox("25년 실적 컬럼", num_cols, index=default_25_idx)
col_26 = st.sidebar.selectbox("26년 실적 컬럼", num_cols, index=default_26_idx)
x_axis = st.sidebar.selectbox("기준 축 (항목 / 고객사)", other_cols if other_cols else df.columns, index=0)

# 6. 상단 실적 비교 지표
total_25 = float(df[col_25].sum())
total_26 = float(df[col_26].sum())
diff_val = total_26 - total_25
diff_rate = (diff_val / total_25 * 100) if total_25 != 0 else 0.0

c1, c2, c3, c4 = st.columns(4)
c1.metric("25년 총 실적", f"{total_25:,.0f}")
c2.metric("26년 총 실적", f"{total_26:,.0f}")
c3.metric("실적 증감액", f"{diff_val:+,.0f}")
c4.metric("증감 퍼센트", f"{diff_rate:+.2f}%")

st.markdown("---")

# 7. 핵심 그래프 2종 (25년 vs 26년 비교 막대 & 비중 차트)
col_left, col_right = st.columns([6, 4])

with col_left:
    st.subheader(f"📌 {x_axis}별 25년 vs 26년 실적 비교")
    chart_data = df.groupby(x_axis, as_index=False)[[col_25, col_26]].sum().sort_values(by=col_26, ascending=False).head(15)
    
    fig_bar = px.bar(
        chart_data,
        x=x_axis,
        y=[col_25, col_26],
        barmode='group',
        labels={"value": "실적금액", "variable": "구분"},
        color_discrete_map={col_25: "#93C5FD", col_26: "#1D4ED8"}
    )
    fig_bar.update_layout(height=450, xaxis_tickangle=-45, template="plotly_white", legend=dict(orientation="h", y=1.1, x=0))
    st.plotly_chart(fig_bar, use_container_width=True)

with col_right:
    st.subheader(f"🥧 26년 {x_axis} 점유율 비중")
    fig_pie = px.pie(
        chart_data.head(8),
        names=x_axis,
        values=col_26,
        hole=0.4
    )
    fig_pie.update_traces(textposition='inside', textinfo='percent+label')
    fig_pie.update_layout(height=450, margin=dict(t=20, b=20, l=10, r=10))
    st.plotly_chart(fig_pie, use_container_width=True)

# 8. 원본 데이터 간략 확인
with st.expander("📄 데이터 테이블 확인"):
    st.dataframe(df, use_container_width=True)
