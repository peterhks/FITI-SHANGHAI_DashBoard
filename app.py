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
        <div style="font-size:18px; font-weight:700;">FITI SHANGHAI 지사 실적 종합 분석</div>
        <div style="font-size:12px; color:#D0E1FD;">상해지사 사업 실적 및 시각화 대시보드</div>
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
    # 파일 로드 실패 시 기본 데모 데이터
    return pd.DataFrame({
        "구분": ["중국 GB시험", "KC인증", "바이어 매뉴얼", "공장 완제품검사", "위생용품"],
        "실적금액": [45000000, 32000000, 28000000, 19000000, 12000000],
        "건수": [150, 110, 95, 60, 40]
    })

# 사이드바 업로드
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
        # 쉼표 제거 후 숫자로 변환 가능한지 체크
        cleaned = df[col].astype(str).str.replace(',', '').str.strip()
        converted = pd.to_numeric(cleaned, errors='coerce')
        if converted.notnull().mean() > 0.6:
            df[col] = converted.fillna(0)

# 숫자형 컬럼과 텍스트/날짜 컬럼 분리
num_cols = df.select_dtypes(include=['number']).columns.tolist()
other_cols = [c for c in df.columns if c not in num_cols]

if not num_cols:
    st.error("데이터에 분석할 수 있는 숫자(실적/금액/건수 등) 컬럼이 없습니다.")
    st.stop()

# 5. 간편 그래프 축 선택
st.sidebar.markdown("---")
st.sidebar.subheader("📊 그래프 축 선택")
x_axis = st.sidebar.selectbox("기준 축 (항목 / 고객사 / 일자)", other_cols if other_cols else df.columns, index=0)
y_axis = st.sidebar.selectbox("실적 지표 (금액 / 건수)", num_cols, index=0)

# 6. 상단 단순 요약 지표
c1, c2, c3 = st.columns(3)
c1.metric(f"총 {y_axis} 합계", f"{df[y_axis].sum():,.0f}")
c2.metric(f"평균 {y_axis}", f"{df[y_axis].mean():,.1f}")
c3.metric("전체 데이터 건수", f"{len(df):,} 건")

st.markdown("---")

# 7. 핵심 그래프 2종 (막대 그래프 & 비중 파이 차트)
col_left, col_right = st.columns([6, 4])

with col_left:
    st.subheader(f"📌 {x_axis}별 {y_axis} 비교")
    # 상위 20개 항목으로 자동 집계
    chart_data = df.groupby(x_axis, as_index=False)[y_axis].sum().sort_values(by=y_axis, ascending=False).head(20)
    fig_bar = px.bar(
        chart_data, 
        x=x_axis, 
        y=y_axis, 
        text_auto=',.0f',
        color=y_axis,
        color_continuous_scale="Blues"
    )
    fig_bar.update_layout(height=450, xaxis_tickangle=-45, template="plotly_white")
    st.plotly_chart(fig_bar, use_container_width=True)

with col_right:
    st.subheader(f"🥧 {x_axis} 점유율 비중")
    fig_pie = px.pie(
        chart_data.head(8), 
        names=x_axis, 
        values=y_axis, 
        hole=0.4
    )
    fig_pie.update_traces(textposition='inside', textinfo='percent+label')
    fig_pie.update_layout(height=450, margin=dict(t=20, b=20, l=10, r=10))
    st.plotly_chart(fig_pie, use_container_width=True)

# 8. 원본 데이터 간략 확인
with st.expander("📄 데이터 테이블 확인"):
    st.dataframe(df, use_container_width=True)
