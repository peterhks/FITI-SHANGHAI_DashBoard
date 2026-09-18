import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
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
        background: linear-gradient(135deg, #002B5C 0%, #003876 100%);
        padding: 22px 28px;
        border-radius: 10px;
        display: flex;
        align-items: center;
        gap: 22px;
        color: #FFFFFF;
        margin-bottom: 22px;
        box-shadow: 0 4px 14px rgba(0, 43, 92, 0.18);
    }
    .fiti-logo-text {
        font-size: 28px;
        font-weight: 900;
        letter-spacing: -0.5px;
        border-right: 1.5px solid rgba(255, 255, 255, 0.25);
        padding-right: 22px;
    }
    .fiti-title-main {
        font-size: 21px;
        font-weight: 800;
        margin-bottom: 4px;
        letter-spacing: -0.3px;
    }
    .fiti-title-sub {
        font-size: 13px;
        color: #D0E1FD;
        font-weight: 400;
    }

    /* 입체형 KPI 카드 스타일 */
    .kpi-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 20px 22px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04), 0 2px 4px -2px rgba(0, 0, 0, 0.04);
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
# 3. 데이터 로드 (항상 '종합' 시트 자동 연결 & 사이드바 UI 제거)
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

df_summary = None

if target_file:
    try:
        if str(target_file).endswith('.csv') or (hasattr(target_file, 'name') and target_file.name.endswith('.csv')):
            df_summary = clean_data(pd.read_csv(target_file))
        else:
            xl = pd.ExcelFile(target_file)
            sheet_names = xl.sheet_names
            # '종합' 시트가 있으면 바로 선택, 없으면 첫 번째 시트 자동 로드
            target_sheet = next((s for s in sheet_names if "종합" in s), sheet_names[0])
            df_summary = clean_data(pd.read_excel(target_file, sheet_name=target_sheet))
    except Exception:
        df_summary = None

if df_summary is None:
    df_summary = pd.DataFrame({
        "사업구분": ["글로벌 바이어", "패션잡화", "GB", "제품평가"],
        "2025년": [2127814670, 1726058860, 1144510000, 343665080],
        "2026년": [2151187420, 1869389240, 1894452100, 490344100]
    })

# =========================================================
# 4. 내부 자동 컬럼 인식 및 4대 사업구분 정제 (사용자 입력 불필요)
# =========================================================
num_cols = df_summary.select_dtypes(include=['number']).columns.tolist()
other_cols = [c for c in df_summary.columns if c not in num_cols]

col_25 = next((c for c in num_cols if "25" in str(c)), num_cols[0] if num_cols else "2025년")
col_26 = next((c for c in num_cols if "26" in str(c)), num_cols[1] if len(num_cols) > 1 else num_cols[0])

# 사업명 텍스트가 들어있는 컬럼 자동 감지
biz_col = other_cols[0] if other_cols else df_summary.columns[0]
for c in other_cols:
    sample_text = "".join(df_summary[c].dropna().astype(str).tolist())
    if any(k in sample_text for k in ["글로벌", "패션", "GB", "제품평가"]):
        biz_col = c
        break

# 4대 카테고리 매핑 함수
def map_biz_category(val):
    s = str(val).replace(" ", "").upper()
    if "글로벌" in s or "GLOBAL" in s or "BUYER" in s:
        return "글로벌 바이어"
    elif "패션" in s or "잡화" in s or "FASHION" in s:
        return "패션잡화"
    elif "GB" in s or "중국" in s:
        return "GB"
    elif "제품평가" in s or "검사" in s or "INSPECTION" in s:
        return "제품평가"
    return None

df_summary["표준사업구분"] = df_summary[biz_col].apply(map_biz_category)

# 4대 사업에 해당하는 순수 데이터만 추출 (TOTAL, SUB TOTAL, 머리글 자동 제거)
target_categories = ["글로벌 바이어", "패션잡화", "GB", "제품평가"]
summary_data = df_summary.dropna(subset=["표준사업구분"]).copy()

# 중복 행이 있을 경우 사업별로 합산 집계
summary_chart = summary_data.groupby("표준사업구분", as_index=False)[[col_25, col_26]].sum()
summary_chart["정렬"] = summary_chart["표준사업구분"].apply(lambda x: target_categories.index(x) if x in target_categories else 99)
summary_chart = summary_chart.sort_values("정렬").reset_index(drop=True)

# =========================================================
# 5. 상단 종합 KPI 카드 (4대 사업 합산)
# =========================================================
total_25 = float(summary_chart[col_25].sum())
total_26 = float(summary_chart[col_26].sum())
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
        <div class="kpi-sub">4대 사업 누적액</div>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown(f"""
    <div class="kpi-card" style="border-top-color: #003876;">
        <div class="kpi-title">🚀 26년 총 실적</div>
        <div class="kpi-num" style="color: #003876;">{total_26:,.0f}</div>
        <div class="kpi-sub">4대 사업 누적액</div>
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
# 6. 사이드바 메뉴: 3가지 카테고리만 심플하게 배치
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
    
    fig_bar = go.Figure()
    
    # 2025년 바 (차분한 Slate Grey)
    fig_bar.add_trace(go.Bar(
        x=summary_chart["표준사업구분"],
        y=summary_chart[col_25],
        name="2025년 총 실적",
        marker=dict(
            color="#94A3B8",
            line=dict(color="#64748B", width=1),
            cornerradius=6
        ),
        text=summary_chart[col_25].apply(lambda x: f"{x/1e8:.1f}억" if x >= 1e8 else f"{x/1e4:.0f}만"),
        textposition="outside",
        textfont=dict(size=12, color="#475569", family="Pretendard")
    ))
    
    # 2026년 바 (세련된 Deep Royal Blue)
    fig_bar.add_trace(go.Bar(
        x=summary_chart["표준사업구분"],
        y=summary_chart[col_26],
        name="2026년 총 실적",
        marker=dict(
            color="#1D4ED8",
            line=dict(color="#1E40AF", width=1),
            cornerradius=6
        ),
        text=summary_chart[col_26].apply(lambda x: f"{x/1e8:.1f}억" if x >= 1e8 else f"{x/1e4:.0f}만"),
        textposition="outside",
        textfont=dict(size=12, color="#0F172A", family="Pretendard", weight="bold")
    ))
    
    fig_bar.update_layout(
        height=450,
        bargap=0.32,
        bargroupgap=0.10,
        yaxis=dict(
            rangemode='tozero',
            title=dict(text="실적금액 (원)", font=dict(size=12, color="#64748B")),
            gridcolor="#F1F5F9",
            zerolinecolor="#E2E8F0"
        ),
        xaxis=dict(
            categoryorder='array',
            categoryarray=target_categories,
            tickfont=dict(size=14, weight="bold", color="#1E293B")
        ),
        template="plotly_white",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.05,
            xanchor="left",
            x=0,
            font=dict(size=12, color="#334155")
        ),
        margin=dict(t=50, b=20, l=10, r=10)
    )
    st.plotly_chart(fig_bar, use_container_width=True)

    st.write("")
    st.subheader("🥧 사업별 점유율 비중 (전체 실적 기준)")
    
    biz_colors = {
        "글로벌 바이어": "#2563EB",
        "패션잡화": "#F59E0B",
        "GB": "#10B981",
        "제품평가": "#8B5CF6"
    }
    
    pie_col1, pie_col2 = st.columns(2)
    
    with pie_col1:
        fig_pie_25 = px.pie(
            summary_chart,
            names="표준사업구분",
            values=col_25,
            hole=0.55,
            title="2025년 사업별 실적 점유율",
            category_orders={"표준사업구분": target_categories},
            color="표준사업구분",
            color_discrete_map=biz_colors
        )
        fig_pie_25.update_traces(
            textposition='inside',
            textinfo='percent+label',
            textfont=dict(size=13, family="Pretendard"),
            marker=dict(line=dict(color='#FFFFFF', width=2))
        )
        fig_pie_25.update_layout(
            height=430,
            margin=dict(t=50, b=20, l=10, r=10),
            legend=dict(orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5)
        )
        st.plotly_chart(fig_pie_25, use_container_width=True)
        
    with pie_col2:
        fig_pie_26 = px.pie(
            summary_chart,
            names="표준사업구분",
            values=col_26,
            hole=0.55,
            title="2026년 사업별 실적 점유율",
            category_orders={"표준사업구분": target_categories},
            color="표준사업구분",
            color_discrete_map=biz_colors
        )
        fig_pie_26.update_traces(
            textposition='inside',
            textinfo='percent+label',
            textfont=dict(size=13, family="Pretendard"),
            marker=dict(line=dict(color='#FFFFFF', width=2))
        )
        fig_pie_26.update_layout(
            height=430,
            margin=dict(t=50, b=20, l=10, r=10),
            legend=dict(orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5)
        )
        st.plotly_chart(fig_pie_26, use_container_width=True)

# [두번째장] 각 사업별 년도 대비 실적 비교
elif page_menu == "두번째장 : 각 사업별 년도 대비 실적 비교":
    st.subheader("🏢 사업별 2025년 vs 2026년 실적 증감 비교")
    
    biz_df = summary_chart.copy()
    biz_df["증감액"] = biz_df[col_26] - biz_df[col_25]
    biz_df["증감률(%)"] = (biz_df["증감액"] / biz_df[col_25].replace(0, pd.NA) * 100).fillna(0)
    biz_df = biz_df.sort_values(by=col_26, ascending=False)
    
    col2_l, col2_r = st.columns([6, 4])
    
    with col2_l:
        biz_chart = biz_df.rename(columns={col_25: "2025년 실적", col_26: "2026년 실적", "표준사업구분": "사업구분"})
        fig2_bar = px.bar(
            biz_chart,
            x="사업구분",
            y=["2025년 실적", "2026년 실적"],
            barmode='group',
            labels={"value": "실적금액 (원)", "variable": "연도 구분"},
            color_discrete_map={"2025년 실적": "#94A3B8", "2026년 실적": "#1D4ED8"}
        )
        fig2_bar.update_layout(
            height=440,
            yaxis=dict(rangemode='tozero', title="실적금액 (원)"),
            template="plotly_white",
            legend=dict(orientation="h", y=1.12, x=0)
        )
        st.plotly_chart(fig2_bar, use_container_width=True)

    with col2_r:
        fig2_diff = px.bar(
            biz_df,
            x="표준사업구분",
            y="증감액",
            text_auto=',.0f',
            title="사업별 실적 증감액 (26년 - 25년)",
            color="증감액",
            color_continuous_scale=["#2563EB", "#CBD5E1", "#E11D48"]
        )
        fig2_diff.update_layout(height=440, template="plotly_white")
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
    st.subheader("🤝 각 사업별 협력사 실적 현황")
    
    unique_biz = summary_chart["표준사업구분"].tolist()
    selected_biz = st.selectbox("조회할 사업부문을 선택하세요:", unique_biz)
    
    st.info(f"선택하신 **[{selected_biz}]** 부문의 상세 협력사별 데이터는 다음 단계에서 각 개별 사업 시트와 정밀 연동될 예정입니다.")
