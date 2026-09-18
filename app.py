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
# 3. 데이터 로드 및 다중 시트 지원
# =========================================================
EXCEL_FILE = "복사본 performance_260825.xlsx"

st.sidebar.markdown("### 📁 데이터 관리")
uploaded_file = st.sidebar.file_uploader("실적 엑셀/CSV 파일 업로드", type=["xlsx", "csv"])

target_file = uploaded_file if uploaded_file else (EXCEL_FILE if os.path.exists(EXCEL_FILE) else None)

def clean_data(data_frame):
    for col in data_frame.columns:
        if data_frame[col].dtype == object:
            cleaned = data_frame[col].astype(str).str.replace(',', '').str.strip()
            converted = pd.to_numeric(cleaned, errors='coerce')
            if converted.notnull().mean() > 0.6:
                data_frame[col] = converted.fillna(0)
    return data_frame

if target_file:
    try:
        if str(target_file).endswith('.csv') or (hasattr(target_file, 'name') and target_file.name.endswith('.csv')):
            df = clean_data(pd.read_csv(target_file))
        else:
            xl = pd.ExcelFile(target_file)
            sheet_names = xl.sheet_names
            
            # 시트가 여러 개일 경우 선택 가능하도록 지원
            if len(sheet_names) > 1:
                selected_sheet = st.sidebar.selectbox("분석할 시트 선택", sheet_names, index=0)
                df = clean_data(pd.read_excel(target_file, sheet_name=selected_sheet))
            else:
                df = clean_data(pd.read_excel(target_file))
    except Exception:
        target_file = None

if not target_file:
    # 예시 샘플 데이터
    df = pd.DataFrame({
        "사업구분": ["중국 GB시험", "KC인증", "바이어 매뉴얼", "공장 완제품검사", "위생용품/기구용기"],
        "바이어명": ["POLO RALPH LAUREN", "F&F", "무신사", "삼성물산", "FILA"],
        "25년 1월": [305000000, 180000000, 120000000, 85000000, 35000000],
        "26년 1월": [362000000, 195000000, 140000000, 92000000, 41000000]
    })

num_cols = df.select_dtypes(include=['number']).columns.tolist()
other_cols = [c for c in df.columns if c not in num_cols]

if not num_cols:
    st.error("데이터에 분석 가능한 숫자(실적 금액) 컬럼이 없습니다.")
    st.stop()

# =========================================================
# 4. 사이드바 컬럼 및 축 설정
# =========================================================
st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ 기준 컬럼 설정")

col_25 = st.sidebar.selectbox("2025년 실적 컬럼", num_cols, index=next((i for i, c in enumerate(num_cols) if "25" in str(c)), 0))
col_26 = st.sidebar.selectbox("2026년 실적 컬럼", num_cols, index=next((i for i, c in enumerate(num_cols) if "26" in str(c)), 1 if len(num_cols) > 1 else 0))

# 업체명이 아닌 '사업구분'과 '바이어명' 분리 선택
default_biz_idx = next((i for i, c in enumerate(other_cols) if any(k in str(c) for k in ["사업", "구분", "분류", "항목", "대분류"])), 0)
default_buyer_idx = next((i for i, c in enumerate(other_cols) if any(k in str(c) for k in ["바이어", "고객", "거래처", "브랜드"])), 1 if len(other_cols) > 1 else 0)

biz_col = st.sidebar.selectbox("사업 구분 기준 컬럼", other_cols if other_cols else df.columns, index=default_biz_idx)
buyer_col = st.sidebar.selectbox("바이어(고객사) 기준 컬럼", other_cols if other_cols else df.columns, index=default_buyer_idx)

# =========================================================
# 5. 상단 종합 KPI 카드 (TOTAL 행 제외 집계)
# =========================================================
calc_df = df[~df[buyer_col].astype(str).str.upper().str.contains("TOTAL|합계|소계", na=False)]
if calc_df.empty:
    calc_df = df

total_25 = float(calc_df[col_25].sum())
total_26 = float(calc_df[col_26].sum())
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
        <span class="kpi-badge" style="background-color: {badge_bg}; color: {diff_color};">전년 대비 실적차</span>
    </div>
    """, unsafe_allow_html=True)

with c4:
    st.markdown(f"""
    <div class="kpi-card" style="border-top-color: {diff_color};">
        <div class="kpi-title">📊 증감 퍼센트</div>
        <div class="kpi-num" style="color: {diff_color};">{diff_sign}{diff_rate:0.2f}%</div>
        <span class="kpi-badge" style="background-color: {badge_bg}; color: {diff_color};">전년 대비 성장률</span>
    </div>
    """, unsafe_allow_html=True)

st.write("")
st.markdown("---")

# =========================================================
# 6. 사이드바 메뉴: 3가지 카테고리
# =========================================================
st.sidebar.markdown("### 📑 분석 페이지 선택")
page_menu = st.sidebar.radio(
    "이동할 카테고리를 선택하세요:",
    [
        "첫번째장 : 종합 실적 현황",
        "두번째장 : 각 사업별 년도 대비 실적 비교",
        "세번째장 : 각 사업별 협력사 비교"
    ],
    index=0
)

# =========================================================
# 7. 본문 페이지 렌더링
# =========================================================

# [첫번째장] 종합 실적 현황
if page_menu == "첫번째장 : 종합 실적 현황":
    st.subheader("📌 2025년 총 실적 vs 2026년 총 실적 비교")
    
    # 1. 상위 바이어 실적 비교 막대 차트
    chart_data = calc_df.groupby(buyer_col, as_index=False)[[col_25, col_26]].sum().sort_values(by=col_26, ascending=False).head(15)
    chart_data_renamed = chart_data.rename(columns={col_25: "2025년 총 실적", col_26: "2026년 총 실적"})
    
    fig1_bar = px.bar(
        chart_data_renamed,
        x=buyer_col,
        y=["2025년 총 실적", "2026년 총 실적"],
        barmode='group',
        labels={"value": "실적금액 (원)", "variable": "실적 구분"},
        color_discrete_map={"2025년 총 실적": "#93C5FD", "2026년 총 실적": "#003876"}
    )
    fig1_bar.update_layout(
        height=430,
        xaxis_tickangle=-45,
        yaxis=dict(rangemode='tozero', title="실적금액 (원)"),
        template="plotly_white",
        legend=dict(orientation="h", y=1.12, x=0)
    )
    st.plotly_chart(fig1_bar, use_container_width=True)

    st.write("")
    # 2. 사업별 점유율 비중 (수십 개 업체명이 아닌, 사업/대분류 단위로 상위 7개 + 기타로 묶어 표기)
    st.subheader("🥧 사업별 점유율 비중 (전체 실적 기준)")
    
    biz_df = calc_df.groupby(biz_col, as_index=False)[[col_25, col_26]].sum()
    
    # 사업이 너무 많을 경우 상위 6개 외에는 '기타'로 합산하여 깔끔하게 정리
    def group_top_categories(data_frame, target_col, val_column, top_n=6):
        sorted_df = data_frame.sort_values(by=val_column, ascending=False)
        if len(sorted_df) > top_n:
            top_part = sorted_df.head(top_n).copy()
            etc_sum = sorted_df.iloc[top_n:][val_column].sum()
            etc_row = pd.DataFrame([{target_col: "기타 사업", val_column: etc_sum}])
            return pd.concat([top_part[[target_col, val_column]], etc_row], ignore_index=True)
        return sorted_df[[target_col, val_column]]

    pie_25_data = group_top_categories(biz_df, biz_col, col_25)
    pie_26_data = group_top_categories(biz_df, biz_col, col_26)
    
    pie_col1, pie_col2 = st.columns(2)
    
    with pie_col1:
        fig_pie_25 = px.pie(
            pie_25_data,
            names=biz_col,
            values=col_25,
            hole=0.45,
            title="2025년 사업별 실적 점유율",
            color_discrete_sequence=px.colors.qualitative.Prism
        )
        fig_pie_25.update_traces(textposition='inside', textinfo='percent+label')
        fig_pie_25.update_layout(height=450, margin=dict(t=50, b=20, l=10, r=10))
        st.plotly_chart(fig_pie_25, use_container_width=True)
        
    with pie_col2:
        fig_pie_26 = px.pie(
            pie_26_data,
            names=biz_col,
            values=col_26,
            hole=0.45,
            title="2026년 사업별 실적 점유율",
            color_discrete_sequence=px.colors.qualitative.Safe
        )
        fig_pie_26.update_traces(textposition='inside', textinfo='percent+label')
        fig_pie_26.update_layout(height=450, margin=dict(t=50, b=20, l=10, r=10))
        st.plotly_chart(fig_pie_26, use_container_width=True)

# [두번째장] 각 사업별 년도 대비 실적 비교
elif page_menu == "두번째장 : 각 사업별 년도 대비 실적 비교":
    st.subheader(f"🏢 {biz_col}별 2025년 vs 2026년 실적 증감 비교")
    
    biz_df = calc_df.groupby(biz_col, as_index=False)[[col_25, col_26]].sum()
    biz_df["증감액"] = biz_df[col_26] - biz_df[col_25]
    biz_df["증감률(%)"] = (biz_df["증감액"] / biz_df[col_25].replace(0, pd.NA) * 100).fillna(0)
    biz_df = biz_df.sort_values(by=col_26, ascending=False)
    
    col2_l, col2_r = st.columns([6, 4])
    
    with col2_l:
        biz_chart = biz_df.rename(columns={col_25: "2025년 실적", col_26: "2026년 실적"})
        fig2_bar = px.bar(
            biz_chart,
            x=biz_col,
            y=["2025년 실적", "2026년 실적"],
            barmode='group',
            labels={"value": "실적금액 (원)", "variable": "연도 구분"},
            color_discrete_map={"2025년 실적": "#CBD5E1", "2026년 실적": "#2563EB"}
        )
        fig2_bar.update_layout(
            height=450,
            xaxis_tickangle=-30,
            yaxis=dict(rangemode='tozero', title="실적금액 (원)"),
            template="plotly_white",
            legend=dict(orientation="h", y=1.12, x=0)
        )
        st.plotly_chart(fig2_bar, use_container_width=True)

    with col2_r:
        fig2_diff = px.bar(
            biz_df,
            x=biz_col,
            y="증감액",
            text_auto=',.0f',
            title="사업별 실적 증감액 (26년 - 25년)",
            color="증감액",
            color_continuous_scale=["#2563EB", "#CBD5E1", "#E11D48"]
        )
        fig2_diff.update_layout(height=450, xaxis_tickangle=-30, template="plotly_white")
        st.plotly_chart(fig2_diff, use_container_width=True)
        
    st.markdown("##### 📋 사업별 세부 실적 요약 테이블")
    st.dataframe(biz_df.style.format({
        col_25: "{:,.0f}",
        col_26: "{:,.0f}",
        "증감액": "{:+,.0f}",
        "증감률(%)": "{:+.2f}%"
    }), use_container_width=True)

# [세번째장] 각 사업별 협력사 비교
elif page_menu == "세번째장 : 각 사업별 협력사 비교":
    st.subheader(f"🤝 각 사업별 {buyer_col}(협력사) 실적 현황")
    
    unique_biz = calc_df[biz_col].dropna().unique().tolist()
    selected_biz = st.selectbox("조회할 사업부문을 선택하세요:", unique_biz)
    
    filtered_df = calc_df[calc_df[biz_col] == selected_biz]
    buyer_group = filtered_df.groupby(buyer_col, as_index=False)[[col_25, col_26]].sum().sort_values(by=col_26, ascending=False).head(15)
    buyer_group["증감액"] = buyer_group[col_26] - buyer_group[col_25]
    buyer_group["증감률(%)"] = (buyer_group["증감액"] / buyer_group[col_25].replace(0, pd.NA) * 100).fillna(0)
    
    col3_l, col3_r = st.columns([6, 4])
    
    with col3_l:
        buyer_chart = buyer_group.rename(columns={col_25: "2025년 실적", col_26: "2026년 실적"})
        fig3_bar = px.bar(
            buyer_chart,
            x=buyer_col,
            y=["2025년 실적", "2026년 실적"],
            barmode='group',
            labels={"value": "실적금액 (원)", "variable": "연도 구분"},
            color_discrete_map={"2025년 실적": "#93C5FD", "2026년 실적": "#003876"}
        )
        fig3_bar.update_layout(
            title=f"[{selected_biz}] 상위 협력사 실적 비교",
            height=450,
            xaxis_tickangle=-45,
            yaxis=dict(rangemode='tozero', title="실적금액 (원)"),
            template="plotly_white",
            legend=dict(orientation="h", y=1.12, x=0)
        )
        st.plotly_chart(fig3_bar, use_container_width=True)
        
    with col3_r:
        fig3_pie = px.pie(
            buyer_group.head(6),
            names=buyer_col,
            values=col_26,
            hole=0.4,
            title=f"[{selected_biz}] 2026년 협력사별 비중"
        )
        fig3_pie.update_traces(textposition='inside', textinfo='percent+label')
        fig3_pie.update_layout(height=450, margin=dict(t=40, b=20, l=10, r=10))
        st.plotly_chart(fig3_pie, use_container_width=True)

    st.markdown(f"##### 📋 [{selected_biz}] 협력사별 실적 상세표")
    st.dataframe(buyer_group.style.format({
        col_25: "{:,.0f}",
        col_26: "{:,.0f}",
        "증감액": "{:+,.0f}",
        "증감률(%)": "{:+.2f}%"
    }), use_container_width=True)
