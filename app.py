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
        <div class="fiti-title-main">상해지사 실적 종합 분석 시스템</div>
        <div class="fiti-title-sub">상해지사 사업 실적 및 분석 시스템 | 상해지사 사업팀</div>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================
# 3. 데이터 로드 및 정제 유틸리티
# =========================================================
EXCEL_FILE = "복사본 performance_260825.xlsx"

st.sidebar.markdown("### 📁 데이터 관리")
uploaded_file = st.sidebar.file_uploader("실적 엑셀 파일 업로드", type=["xlsx", "csv"])
target_file = uploaded_file if uploaded_file else (EXCEL_FILE if os.path.exists(EXCEL_FILE) else None)

def clean_data(data_frame):
    for col in data_frame.columns:
        if data_frame[col].dtype == object:
            cleaned = data_frame[col].astype(str).str.replace(',', '').str.strip()
            converted = pd.to_numeric(cleaned, errors='coerce')
            if converted.notnull().mean() > 0.6:
                data_frame[col] = converted.fillna(0)
    return data_frame

excel_obj = None
sheet_dict = {}

if target_file:
    try:
        excel_obj = pd.ExcelFile(target_file)
        for s in excel_obj.sheet_names:
            sheet_dict[s.strip().lower()] = s
    except Exception as e:
        st.error(f"엑셀 파일 로드 실패: {e}")
        st.stop()
else:
    st.warning("분석할 엑셀 파일을 업로드해 주세요.")
    st.stop()

def get_sheet_by_keyword(keywords):
    for s_clean, orig_name in sheet_dict.items():
        if all(k.lower() in s_clean for k in keywords):
            return orig_name
    return None

# =========================================================
# 4. '종합' 시트 정밀 파싱 (병합 셀 ffill 처리 적용)
# =========================================================
summary_sheet_name = get_sheet_by_keyword(["종합"]) or excel_obj.sheet_names[0]
raw_summary = pd.read_excel(target_file, sheet_name=summary_sheet_name)

header_row_idx = 0
for idx, row in raw_summary.iterrows():
    row_text = "".join(row.dropna().astype(str).tolist())
    if "구분" in row_text and ("25" in row_text or "합계" in row_text):
        header_row_idx = idx
        break

df_summary = pd.read_excel(target_file, sheet_name=summary_sheet_name, skiprows=header_row_idx)
df_summary = clean_data(df_summary)

num_cols = df_summary.select_dtypes(include=['number']).columns.tolist()
col_25 = next((c for c in num_cols if "25" in str(c)), num_cols[0] if num_cols else None)
col_26 = next((c for c in num_cols if "26" in str(c)), num_cols[1] if len(num_cols) > 1 else num_cols[0])

other_cols = [c for c in df_summary.columns if c not in num_cols]
cat_col = other_cols[0] if other_cols else df_summary.columns[0]
sub_cat_col = other_cols[1] if len(other_cols) > 1 else None

for c in other_cols:
    sample_str = "".join(df_summary[c].dropna().astype(str).tolist())
    if any(k in sample_str for k in ["패션잡화", "GB", "글로벌", "제품평가"]):
        cat_col = c
        break

df_summary[cat_col] = df_summary[cat_col].replace(r'^\s*$', pd.NA, regex=True)
df_summary["사업구분_채움"] = df_summary[cat_col].ffill()

def map_biz_category(val):
    s = str(val).replace(" ", "").upper()
    if "글로벌" in s or "GLOBAL" in s:
        return "글로벌 바이어"
    elif "패션" in s or "잡화" in s or "KC" in s:
        return "패션잡화"
    elif "GB" in s:
        return "GB"
    elif "제품평가" in s or "INSPECTION" in s:
        return "제품평가"
    return None

df_summary["표준사업구분"] = df_summary["사업구분_채움"].apply(map_biz_category)

exclude_pattern = r"SUB\s*TOTAL|TOTAL|합계|소계"
calc_summary = df_summary[
    (~df_summary[cat_col].astype(str).str.strip().str.upper().str.contains(exclude_pattern, regex=True, na=False)) &
    (df_summary["표준사업구분"].notnull())
].copy()

if sub_cat_col:
    calc_summary["세부항목"] = calc_summary[sub_cat_col].fillna(calc_summary["표준사업구분"]).astype(str)
else:
    calc_summary["세부항목"] = calc_summary["표준사업구분"]

target_categories = ["글로벌 바이어", "패션잡화", "GB", "제품평가"]
summary_chart = calc_summary.groupby("표준사업구분", as_index=False)[[col_25, col_26]].sum()
summary_chart["정렬"] = summary_chart["표준사업구분"].apply(lambda x: target_categories.index(x) if x in target_categories else 99)
summary_chart = summary_chart.sort_values("정렬").reset_index(drop=True)

# =========================================================
# 5. 세부 파트 데이터 로드 및 검증
# =========================================================
PART_SHEET_MAPPINGS = {
    "패션잡화": [["kc"]],
    "GB": [["gb"]],
    "글로벌 바이어": [["global", "1"], ["global", "2"]],
    "제품평가": [["inspection", "원단"], ["inspection", "가먼트"]]
}

def load_part_data(categories_target):
    keywords_list = PART_SHEET_MAPPINGS.get(categories_target, [])
    matched_dfs = []
    
    for k_words in keywords_list:
        actual_sheet = get_sheet_by_keyword(k_words)
        if actual_sheet:
            try:
                temp_df = clean_data(pd.read_excel(target_file, sheet_name=actual_sheet))
                b_nums = temp_df.select_dtypes(include=['number']).columns.tolist()
                b_others = [c for c in temp_df.columns if c not in b_nums]
                
                if b_nums and b_others:
                    b_25 = next((c for c in b_nums if "25" in str(c)), b_nums[0])
                    b_26 = next((c for c in b_nums if "26" in str(c)), b_nums[1] if len(b_nums) > 1 else b_nums[0])
                    b_name = next((c for c in b_others if any(k in str(c) for k in ["바이어", "고객", "업체", "거래처", "브랜드"])), b_others[0])
                    
                    exclude_pat = r"TOTAL|SUB\s*TOTAL|합계|소계|누계|^구분$|상해지사\s*사업코드"
                    temp_df = temp_df[
                        (~temp_df[b_name].astype(str).str.strip().str.upper().str.contains(exclude_pat, regex=True, na=False)) &
                        (temp_df[b_name].notnull()) &
                        (temp_df[b_name].astype(str).str.strip() != "")
                    ].copy()
                    
                    temp_df = temp_df[[b_name, b_25, b_26]].rename(columns={
                        b_name: "바이어명",
                        b_25: "2025년 실적",
                        b_26: "2026년 실적"
                    })
                    matched_dfs.append(temp_df)
            except Exception:
                pass
                
    if matched_dfs:
        combined = pd.concat(matched_dfs, ignore_index=True)
        return combined.groupby("바이어명", as_index=False)[["2025년 실적", "2026년 실적"]].sum()
    return pd.DataFrame(columns=["바이어명", "2025년 실적", "2026년 실적"])

audit_results = []
part_data_cache = {}

for cat in target_categories:
    sum_row = summary_chart[summary_chart["표준사업구분"] == cat]
    sum_25 = float(sum_row[col_25].values[0]) if not sum_row.empty else 0.0
    sum_26 = float(sum_row[col_26].values[0]) if not sum_row.empty else 0.0
    
    p_df = load_part_data(cat)
    part_data_cache[cat] = p_df
    part_25 = float(p_df["2025년 실적"].sum()) if not p_df.empty else 0.0
    part_26 = float(p_df["2026년 실적"].sum()) if not p_df.empty else 0.0
    
    diff_25 = abs(sum_25 - part_25)
    diff_26 = abs(sum_26 - part_26)
    is_match = (diff_25 < 100) and (diff_26 < 100)
    
    audit_results.append({
        "사업구분": cat,
        "종합 실적 (25년)": sum_25,
        "파트 실적 (25년)": part_25,
        "종합 실적 (26년)": sum_26,
        "파트 실적 (26년)": part_26,
        "25년 오차": sum_25 - part_25,
        "26년 오차": sum_26 - part_26,
        "일치여부": "✅ 일치 (정상)" if is_match else "⚠️ 불일치 확인 필요"
    })

audit_df = pd.DataFrame(audit_results)

# =========================================================
# 6. 사이드바 메뉴
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
# 7. 상단 종합 KPI 카드 (두번째장 & 세번째장 모두 사업별 선택과 연동)
# =========================================================
selected_view_for_card = "전체 사업 보기"

if page_menu == "두번째장 : 각 사업별 년도 대비 실적 비교":
    if "selected_biz_view" not in st.session_state:
        st.session_state["selected_biz_view"] = "전체 사업 보기"
    selected_view_for_card = st.session_state["selected_biz_view"]
elif page_menu == "세번째장 : 각 사업별 협력사 비교":
    if "selected_tab3_biz" not in st.session_state:
        st.session_state["selected_tab3_biz"] = target_categories[0]
    selected_view_for_card = st.session_state["selected_tab3_biz"]

if selected_view_for_card != "전체 사업 보기" and selected_view_for_card in target_categories:
    target_row = summary_chart[summary_chart["표준사업구분"] == selected_view_for_card]
    total_25 = float(target_row[col_25].sum()) if not target_row.empty else 0.0
    total_26 = float(target_row[col_26].sum()) if not target_row.empty else 0.0
    card_sub_desc = f"[{selected_view_for_card}] 실적 합계"
else:
    total_25 = float(summary_chart[col_25].sum())
    total_26 = float(summary_chart[col_26].sum())
    card_sub_desc = "종합 TOTAL 합계 (정상 일치)"

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
        <div class="kpi-sub">{card_sub_desc}</div>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown(f"""
    <div class="kpi-card" style="border-top-color: #003876;">
        <div class="kpi-title">🚀 26년 총 실적</div>
        <div class="kpi-num" style="color: #003876;">{total_26:,.0f}</div>
        <div class="kpi-sub">{card_sub_desc}</div>
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
# 8. 본문 페이지 렌더링
# =========================================================

# [첫번째장] 종합 실적 현황
if page_menu == "첫번째장 : 종합 실적 현황":
    st.subheader("📌 2025년 총 실적 vs 2026년 총 실적 비교")
    
    fig_bar = go.Figure()
    fig_bar.add_trace(go.Bar(
        x=summary_chart["표준사업구분"],
        y=summary_chart[col_25],
        name="2025년 총 실적",
        marker=dict(color="#94A3B8", line=dict(color="#64748B", width=1), cornerradius=6),
        text=summary_chart[col_25].apply(lambda x: f"{x/1e8:.1f}억" if x >= 1e8 else f"{x/1e4:.0f}만"),
        textposition="outside",
        textfont=dict(size=12, color="#475569", family="Pretendard")
    ))
    fig_bar.add_trace(go.Bar(
        x=summary_chart["표준사업구분"],
        y=summary_chart[col_26],
        name="2026년 총 실적",
        marker=dict(color="#1D4ED8", line=dict(color="#1E40AF", width=1), cornerradius=6),
        text=summary_chart[col_26].apply(lambda x: f"{x/1e8:.1f}억" if x >= 1e8 else f"{x/1e4:.0f}만"),
        textposition="outside",
        textfont=dict(size=12, color="#0F172A", family="Pretendard", weight="bold")
    ))
    fig_bar.update_layout(
        height=450,
        bargap=0.32,
        bargroupgap=0.10,
        yaxis=dict(rangemode='tozero', title=dict(text="실적금액 (원)", font=dict(size=12, color="#64748B")), gridcolor="#F1F5F9", zerolinecolor="#E2E8F0"),
        xaxis=dict(categoryorder='array', categoryarray=target_categories, tickfont=dict(size=14, weight="bold", color="#1E293B")),
        template="plotly_white",
        legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="left", x=0, font=dict(size=12, color="#334155")),
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
            summary_chart, names="표준사업구분", values=col_25, hole=0.55,
            title="2025년 사업별 실적 점유율", category_orders={"표준사업구분": target_categories},
            color="표준사업구분", color_discrete_map=biz_colors
        )
        fig_pie_25.update_traces(textposition='inside', textinfo='percent+label', textfont=dict(size=13, family="Pretendard"), marker=dict(line=dict(color='#FFFFFF', width=2)))
        fig_pie_25.update_layout(height=430, margin=dict(t=50, b=20, l=10, r=10), legend=dict(orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5))
        st.plotly_chart(fig_pie_25, use_container_width=True)
        
    with pie_col2:
        fig_pie_26 = px.pie(
            summary_chart, names="표준사업구분", values=col_26, hole=0.55,
            title="2026년 사업별 실적 점유율", category_orders={"표준사업구분": target_categories},
            color="표준사업구분", color_discrete_map=biz_colors
        )
        fig_pie_26.update_traces(textposition='inside', textinfo='percent+label', textfont=dict(size=13, family="Pretendard"), marker=dict(line=dict(color='#FFFFFF', width=2)))
        fig_pie_26.update_layout(height=430, margin=dict(t=50, b=20, l=10, r=10), legend=dict(orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5))
        st.plotly_chart(fig_pie_26, use_container_width=True)

# [두번째장] 각 사업별 년도 대비 실적 비교
elif page_menu == "두번째장 : 각 사업별 년도 대비 실적 비교":
    st.subheader("🏢 사업별 2025년 vs 2026년 실적 증감 비교")
    
    biz_filter_options = ["전체 사업 보기"] + target_categories
    current_idx = biz_filter_options.index(st.session_state.get("selected_biz_view", "전체 사업 보기"))
    
    selected_view = st.selectbox(
        "조회할 사업부문을 선택하세요:", 
        biz_filter_options, 
        index=current_idx,
        key="selected_biz_selectbox"
    )
    
    if selected_view != st.session_state.get("selected_biz_view"):
        st.session_state["selected_biz_view"] = selected_view
        st.rerun()

    if selected_view == "전체 사업 보기":
        display_df = summary_chart.copy()
        x_col = "표준사업구분"
        x_categories = target_categories
        sub_title_diff = "사업별 실적 증감액 (26년 - 25년)"
    else:
        display_df = calc_summary[calc_summary["표준사업구분"] == selected_view].copy()
        x_col = "세부항목"
        x_categories = display_df["세부항목"].unique().tolist()
        sub_title_diff = f"[{selected_view}] 세부항목별 실적 증감액 (26년 - 25년)"
        
    display_df["증감액"] = display_df[col_26] - display_df[col_25]
    display_df["증감률"] = ((display_df["증감액"] / display_df[col_25].replace(0, pd.NA)) * 100).fillna(0.0)
    
    col2_l, col2_r = st.columns([6, 4])
    
    with col2_l:
        fig2_bar = go.Figure()
        fig2_bar.add_trace(go.Bar(
            x=display_df[x_col],
            y=display_df[col_25],
            name="2025년 실적",
            marker=dict(color="#94A3B8", line=dict(color="#64748B", width=1), cornerradius=6),
            text=display_df[col_25].apply(lambda x: f"{x/1e8:.1f}억" if x >= 1e8 else f"{x/1e4:.0f}만"),
            textposition="outside",
            textfont=dict(size=11, color="#475569", family="Pretendard")
        ))
        fig2_bar.add_trace(go.Bar(
            x=display_df[x_col],
            y=display_df[col_26],
            name="2026년 실적",
            marker=dict(color="#1D4ED8", line=dict(color="#1E40AF", width=1), cornerradius=6),
            text=display_df[col_26].apply(lambda x: f"{x/1e8:.1f}억" if x >= 1e8 else f"{x/1e4:.0f}만"),
            textposition="outside",
            textfont=dict(size=11, color="#0F172A", family="Pretendard", weight="bold")
        ))
        fig2_bar.update_layout(
            height=440,
            bargap=0.30,
            bargroupgap=0.10,
            yaxis=dict(rangemode='tozero', title=dict(text="실적금액 (원)", font=dict(size=12, color="#64748B")), gridcolor="#F1F5F9"),
            xaxis=dict(categoryorder='array', categoryarray=x_categories, tickfont=dict(size=13, weight="bold", color="#1E293B")),
            template="plotly_white",
            legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="left", x=0),
            margin=dict(t=50, b=20, l=10, r=10)
        )
        st.plotly_chart(fig2_bar, use_container_width=True)

    with col2_r:
        diff_colors = ["#E11D48" if v >= 0 else "#2563EB" for v in display_df["증감액"]]
        diff_texts = [f"{'+' if v >= 0 else ''}{v/1e8:.2f}억" if abs(v) >= 1e8 else f"{'+' if v >= 0 else ''}{v/1e4:.0f}만" for v in display_df["증감액"]]
        
        fig2_diff = go.Figure()
        fig2_diff.add_trace(go.Bar(
            x=display_df[x_col],
            y=display_df["증감액"],
            marker=dict(color=diff_colors, cornerradius=6),
            text=diff_texts,
            textposition="outside",
            textfont=dict(size=11, weight="bold", family="Pretendard")
        ))
        fig2_diff.update_layout(
            title=sub_title_diff,
            height=440,
            bargap=0.45,
            yaxis=dict(title=dict(text="증감액 (원)", font=dict(size=12, color="#64748B")), gridcolor="#F1F5F9", zerolinecolor="#CBD5E1"),
            xaxis=dict(categoryorder='array', categoryarray=x_categories, tickfont=dict(size=13, weight="bold", color="#1E293B")),
            template="plotly_white",
            margin=dict(t=50, b=20, l=10, r=10)
        )
        st.plotly_chart(fig2_diff, use_container_width=True)
        
    st.markdown(f"##### 📋 [{'전체 사업' if selected_view == '전체 사업 보기' else selected_view}] 실적 요약 테이블")
    
    table_df = pd.DataFrame({
        "구분": display_df[x_col],
        "2025년 실적 (원)": display_df[col_25],
        "2026년 실적 (원)": display_df[col_26],
        "증감액 (원)": display_df["증감액"],
        "증감률(%)": display_df["증감률"]
    })
    
    st.dataframe(
        table_df,
        column_config={
            "구분": st.column_config.TextColumn("구분", width="medium"),
            "2025년 실적 (원)": st.column_config.NumberColumn("2025년 실적 (원)", format="₩%,d"),
            "2026년 실적 (원)": st.column_config.NumberColumn("2026년 실적 (원)", format="₩%,d"),
            "증감액 (원)": st.column_config.NumberColumn("증감액 (원)", format="₩%+d"),
            "증감률(%)": st.column_config.NumberColumn("증감률(%)", format="%+.2f%%"),
        },
        hide_index=True,
        use_container_width=True
    )
    
    st.markdown("---")
    st.subheader("🔍 종합 시트 vs 세부 파트 시트 실적 일치 검증")
    st.caption("※ 패션잡화=KC part | GB=GB part | 글로벌바이어=global part1+2 | 제품평가=inspection (원단+가먼트)")
    
    st.dataframe(
        audit_df,
        column_config={
            "사업구분": st.column_config.TextColumn("사업구분", width="medium"),
            "종합 실적 (25년)": st.column_config.NumberColumn("종합 (25년)", format="₩%,d"),
            "파트 실적 (25년)": st.column_config.NumberColumn("파트합산 (25년)", format="₩%,d"),
            "종합 실적 (26년)": st.column_config.NumberColumn("종합 (26년)", format="₩%,d"),
            "파트 실적 (26년)": st.column_config.NumberColumn("파트합산 (26년)", format="₩%,d"),
            "25년 오차": st.column_config.NumberColumn("25년 오차", format="₩%+d"),
            "26년 오차": st.column_config.NumberColumn("26년 오차", format="₩%+d"),
            "일치여부": st.column_config.TextColumn("일치 상태", width="medium"),
        },
        hide_index=True,
        use_container_width=True
    )

# [세번째장] 각 사업별 협력사(바이어) 비교 (상위 6개 + 기타 통합 집계)
elif page_menu == "세번째장 : 각 사업별 협력사 비교":
    st.subheader("🤝 각 사업별 주요 바이어 2025년 vs 2026년 실적 변화")
    
    # 세 번째 장 전용 사업부문 셀렉트박스 (상단 카드와 실시간 동기화)
    current_tab3_biz = st.session_state.get("selected_tab3_biz", target_categories[0])
    current_idx3 = target_categories.index(current_tab3_biz) if current_tab3_biz in target_categories else 0
    
    selected_biz = st.selectbox(
        "조회할 사업부문을 선택하세요:", 
        target_categories, 
        index=current_idx3,
        key="tab3_biz_selectbox"
    )
    
    if selected_biz != st.session_state.get("selected_tab3_biz"):
        st.session_state["selected_tab3_biz"] = selected_biz
        st.rerun()
    
    raw_b_chart = part_data_cache.get(selected_biz, pd.DataFrame()).copy()
    
    if raw_b_chart.empty:
        st.warning(f"선택하신 [{selected_biz}] 부문에 해당하는 세부 파트 시트의 데이터를 찾을 수 없습니다.")
    else:
        # -------------------------------------------------------------
        # 상위 6개 바이어 + 기타("-" 및 7위 이하 포함) 집계 로직
        # -------------------------------------------------------------
        # '-' 하이픈, 공란, 결측치인 바이어는 바로 '기타' 대상으로 분류
        is_dash = raw_b_chart["바이어명"].astype(str).str.strip().isin(["-", "–", "—", "", "NAN", "NONE"])
        
        valid_buyers = raw_b_chart[~is_dash].copy()
        dash_buyers = raw_b_chart[is_dash].copy()
        
        # 2026년 실적 기준 상위 6개 추출
        valid_buyers = valid_buyers.sort_values(by="2026년 실적", ascending=False).reset_index(drop=True)
        
        if len(valid_buyers) > 6:
            top6 = valid_buyers.iloc[:6].copy()
            rest = valid_buyers.iloc[6:].copy()
            
            etc_25 = rest["2025년 실적"].sum() + dash_buyers["2025년 실적"].sum()
            etc_26 = rest["2026년 실적"].sum() + dash_buyers["2026년 실적"].sum()
            
            etc_row = pd.DataFrame([{
                "바이어명": "기타",
                "2025년 실적": etc_25,
                "2026년 실적": etc_26
            }])
            top_buyers = pd.concat([top6, etc_row], ignore_index=True)
        else:
            if not dash_buyers.empty:
                etc_row = pd.DataFrame([{
                    "바이어명": "기타",
                    "2025년 실적": dash_buyers["2025년 실적"].sum(),
                    "2026년 실적": dash_buyers["2026년 실적"].sum()
                }])
                top_buyers = pd.concat([valid_buyers, etc_row], ignore_index=True)
            else:
                top_buyers = valid_buyers.copy()

        top_buyers["증감액"] = top_buyers["2026년 실적"] - top_buyers["2025년 실적"]
        top_buyers["증감률"] = ((top_buyers["증감액"] / top_buyers["2025년 실적"].replace(0, pd.NA)) * 100).fillna(0.0)

        # X축 순서: 상위 6개사 + '기타'를 맨 뒤로 배치
        x_buyer_names = [b for b in top_buyers["바이어명"] if b != "기타"] + (["기타"] if "기타" in top_buyers["바이어명"].values else [])

        col3_l, col3_r = st.columns([6, 4])
        
        with col3_l:
            fig3_bar = go.Figure()
            fig3_bar.add_trace(go.Bar(
                x=top_buyers["바이어명"],
                y=top_buyers["2025년 실적"],
                name="2025년 실적",
                marker=dict(color="#94A3B8", line=dict(color="#64748B", width=1), cornerradius=6),
                text=top_buyers["2025년 실적"].apply(lambda x: f"{x/1e8:.1f}억" if x >= 1e8 else (f"{x/1e4:.0f}만" if x > 0 else "0만")),
                textposition="outside",
                textfont=dict(size=10, color="#475569", family="Pretendard")
            ))
            fig3_bar.add_trace(go.Bar(
                x=top_buyers["바이어명"],
                y=top_buyers["2026년 실적"],
                name="2026년 실적",
                marker=dict(color="#1D4ED8", line=dict(color="#1E40AF", width=1), cornerradius=6),
                text=top_buyers["2026년 실적"].apply(lambda x: f"{x/1e8:.1f}억" if x >= 1e8 else (f"{x/1e4:.0f}만" if x > 0 else "0만")),
                textposition="outside",
                textfont=dict(size=10, color="#0F172A", family="Pretendard", weight="bold")
            ))
            fig3_bar.update_layout(
                height=440,
                bargap=0.30,
                bargroupgap=0.10,
                yaxis=dict(rangemode='tozero', title=dict(text="실적금액 (원)", font=dict(size=12, color="#64748B")), gridcolor="#F1F5F9"),
                xaxis=dict(categoryorder='array', categoryarray=x_buyer_names, tickangle=-30, tickfont=dict(size=12, weight="bold", color="#1E293B")),
                template="plotly_white",
                legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="left", x=0),
                margin=dict(t=50, b=20, l=10, r=10)
            )
            st.plotly_chart(fig3_bar, use_container_width=True)

        with col3_r:
            diff_colors = ["#E11D48" if v >= 0 else "#2563EB" for v in top_buyers["증감액"]]
            diff_texts = [f"{'+' if v >= 0 else ''}{v/1e8:.2f}억" if abs(v) >= 1e8 else f"{'+' if v >= 0 else ''}{v/1e4:.0f}만" for v in top_buyers["증감액"]]
            
            fig3_diff = go.Figure()
            fig3_diff.add_trace(go.Bar(
                x=top_buyers["바이어명"],
                y=top_buyers["증감액"],
                marker=dict(color=diff_colors, cornerradius=6),
                text=diff_texts,
                textposition="outside",
                textfont=dict(size=10, weight="bold", family="Pretendard")
            ))
            fig3_diff.update_layout(
                title=f"[{selected_biz}] 바이어별 증감액 (26년 - 25년)",
                height=440,
                bargap=0.45,
                yaxis=dict(title=dict(text="증감액 (원)", font=dict(size=12, color="#64748B")), gridcolor="#F1F5F9", zerolinecolor="#CBD5E1"),
                xaxis=dict(categoryorder='array', categoryarray=x_buyer_names, tickangle=-30, tickfont=dict(size=12, weight="bold", color="#1E293B")),
                template="plotly_white",
                margin=dict(t=50, b=20, l=10, r=10)
            )
            st.plotly_chart(fig3_diff, use_container_width=True)

        st.markdown(f"##### 📋 [{selected_biz}] 상위 6개 바이어 및 기타 실적 요약표")
        
        table_buyer_df = pd.DataFrame({
            "바이어명": top_buyers["바이어명"],
            "2025년 실적 (원)": top_buyers["2025년 실적"],
            "2026년 실적 (원)": top_buyers["2026년 실적"],
            "증감액 (원)": top_buyers["증감액"],
            "증감률(%)": top_buyers["증감률"]
        })
        
        st.dataframe(
            table_buyer_df,
            column_config={
                "바이어명": st.column_config.TextColumn("바이어명", width="medium"),
                "2025년 실적 (원)": st.column_config.NumberColumn("2025년 실적 (원)", format="₩%,d"),
                "2026년 실적 (원)": st.column_config.NumberColumn("2026년 실적 (원)", format="₩%,d"),
                "증감액 (원)": st.column_config.NumberColumn("증감액 (원)", format="₩%+d"),
                "증감률(%)": st.column_config.NumberColumn("증감률(%)", format="%+.2f%%"),
            },
            hide_index=True,
            use_container_width=True
        )
