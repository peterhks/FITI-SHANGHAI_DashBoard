import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# 1. 브라우저 탭 및 레이아웃 설정
st.set_page_config(page_title="상해지사 실적 대시보드", layout="wide", page_icon="📈")

DEFAULT_FILE_PATH = "복사본 performance_260825.xlsx"

# 2. 사이드바 - 실시간 파일 업로드 영역
st.sidebar.title("📌 상해지사 실적 분석")
st.sidebar.markdown("---")

uploaded_file = st.sidebar.file_uploader(
    "📂 최신 엑셀 파일 업로드", 
    type=["xlsx", "xls"],
    help="새로 업데이트된 엑셀 파일을 여기에 끌어다 놓으면 즉시 반영됩니다."
)

# 업로드된 파일이 있으면 그것을 쓰고, 없으면 기본 로컬 파일 사용
if uploaded_file is not None:
    source_file = uploaded_file
    st.sidebar.success(f"✔️ 업로드 파일 반영됨:\n{uploaded_file.name}")
else:
    if os.path.exists(DEFAULT_FILE_PATH):
        source_file = DEFAULT_FILE_PATH
        st.sidebar.info(f"📁 기본 파일 사용 중:\n{DEFAULT_FILE_PATH}")
    else:
        st.sidebar.error(f"기본 파일을 찾을 수 없습니다: {DEFAULT_FILE_PATH}")
        st.stop()

st.sidebar.markdown("---")

# 3. 데이터 로딩 및 전처리 함수
@st.cache_data
def get_sheet_names(file):
    xls = pd.ExcelFile(file)
    return xls.sheet_names

@st.cache_data
def load_sheet_data(file, sheet_name):
    return pd.read_excel(file, sheet_name=sheet_name)

@st.cache_data
def load_summary_data(file):
    try:
        df_raw = pd.read_excel(file, sheet_name='종합', skiprows=1)
        df_raw.columns = ['구분', '세부구분', '매출_25년', '매출_26년', '증감수수료', '증감율']
        df_summary = df_raw.dropna(subset=['매출_25년']).copy()
        return df_summary
    except Exception as e:
        st.error(f"'종합' 시트를 처리하는 중 오류가 발생했습니다: {e}")
        return pd.DataFrame()

# 데이터 로드
try:
    sheet_names = get_sheet_names(source_file)
    df_summary = load_summary_data(source_file)
except Exception as e:
    st.error(f"엑셀 파일 로드 실패: {e}")
    st.stop()

# 4. 사이드바 메뉴 선택
menu = st.sidebar.radio("화면 메뉴 선택", ["1. 전체 경영 실적 요약", "2. 부문별 세부 분석", "3. 데이터 원본 조회"])

# -------------------------------------------------------------
# 메뉴 1: 전체 실적 요약
# -------------------------------------------------------------
if menu == "1. 전체 경영 실적 요약":
    st.title("🏢 상해지사 시험 및 제품평가 사업 실적 대시보드")
    st.caption("2025년 대비 2026년 누적 실적 및 부문별 포트폴리오 요약")
    st.markdown("---")
    
    if not df_summary.empty and 'TOTAL' in df_summary['구분'].values:
        total_data = df_summary[df_summary['구분'] == 'TOTAL'].iloc[0]
        rev_25 = int(total_data['매출_25년'])
        rev_26 = int(total_data['매출_26년'])
        growth = float(total_data['증감율']) * 100
        diff = int(total_data['증감수수료'])
        
        # 최상단 핵심 KPI 카드
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("2025년 총 매출", f"₩ {rev_25:,.0f}")
        kpi2.metric("2026년 총 매출", f"₩ {rev_26:,.0f}")
        kpi3.metric("전년 대비 증감액", f"₩ {diff:+,.0f}", delta=f"{diff:+,.0f}")
        kpi4.metric("전체 성장률", f"{growth:+.2f}%", delta=f"{growth:+.2f}%")
        
        st.markdown("---")
        
        # 부문별 비교 (SUB TOTAL, TOTAL 제외)
        main_dept = df_summary[~df_summary['구분'].isin(['SUB TOTAL', 'TOTAL'])].copy()
        main_dept['부문'] = main_dept['구분'].fillna('') + ' ' + main_dept['세부구분'].fillna('')
        main_dept['부문'] = main_dept['부문'].str.strip()
        
        col1, col2 = st.columns([6, 4])
        
        with col1:
            st.subheader("📊 부문별 25년 vs 26년 매출 비교")
            fig_bar = go.Figure(data=[
                go.Bar(name='2025년', x=main_dept['부문'], y=main_dept['매출_25년'], marker_color='#2b5c8f'),
                go.Bar(name='2026년', x=main_dept['부문'], y=main_dept['매출_26년'], marker_color='#2ca02c')
            ])
            fig_bar.update_layout(
                barmode='group', 
                height=420, 
                margin=dict(l=10, r=10, t=30, b=10),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig_bar, use_container_width=True)
            
        with col2:
            st.subheader("🥧 2026년 사업 부문별 비중")
            fig_pie = px.pie(main_dept, values='매출_26년', names='부문', hole=0.4,
                             color_discrete_sequence=px.colors.qualitative.Safe)
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            fig_pie.update_layout(height=420, margin=dict(l=10, r=10, t=30, b=10), showlegend=False)
            st.plotly_chart(fig_pie, use_container_width=True)
    else:
        st.warning("'종합' 시트의 데이터를 읽을 수 없습니다.")

# -------------------------------------------------------------
# 메뉴 2: 부문별 세부 분석
# -------------------------------------------------------------
elif menu == "2. 부문별 세부 분석":
    st.title("🔍 부문별 고객사 / 벤더 실적 분석")
    
    # 엑셀 내 존재하는 시트 중 상세 분석 대상 필터링
    candidate_sheets = ['global part1 ', 'global part2', 'KC part', 'GB part', 'inspection (원단)', 'inspection (가먼트)']
    available_sheets = [s for s in candidate_sheets if s in sheet_names]
    
    if not available_sheets:
        available_sheets = sheet_names
        
    selected_sheet = st.selectbox("분석할 사업 부문 선택", available_sheets)
    df_part = load_sheet_data(source_file, selected_sheet)
    
    if '26년 합계' in df_part.columns and '업체명' in df_part.columns:
        top_n = st.slider("표시할 상위 업체 수", min_value=5, max_value=25, value=10)
        top_df = df_part.dropna(subset=['26년 합계']).sort_values(by='26년 합계', ascending=False).head(top_n)
        
        fig_rank = px.bar(
            top_df,
            x='26년 합계',
            y='업체명',
            orientation='h',
            text='26년 합계',
            color='증감율' if '증감율' in top_df.columns else None,
            color_continuous_scale='Bluered',
            labels={'26년 합계': '2026년 실적 (KRW)', '업체명': '업체명'},
            title=f"[{selected_sheet}] 실적 상위 Top {top_n} 업체"
        )
        fig_rank.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
        fig_rank.update_layout(yaxis={'categoryorder':'total ascending'}, height=520)
        st.plotly_chart(fig_rank, use_container_width=True)
    else:
        st.info("해당 시트에는 '업체명' 또는 '26년 합계' 컬럼이 없어 차트를 생성하지 않고 테이블로 표시합니다.")
        st.dataframe(df_part, use_container_width=True)

# -------------------------------------------------------------
# 메뉴 3: 데이터 원본 조회
# -------------------------------------------------------------
elif menu == "3. 데이터 원본 조회":
    st.title("📋 엑셀 시트별 원본 데이터 조회")
    target_sheet = st.selectbox("조회할 시트 선택", sheet_names)
    df_show = load_sheet_data(source_file, target_sheet)
    st.dataframe(df_show, use_container_width=True)