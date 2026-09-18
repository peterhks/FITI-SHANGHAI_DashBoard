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
# 2. 상단 공식 배너
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
# 3. 데이터 로드 및 정제 엔진
# =========================================================
EXCEL_FILE = "복사본 performance_260825.xlsx"

st.sidebar.markdown("### 📁 데이터 관리")
uploaded_file = st.sidebar.file_uploader("실적 엑셀 파일 업로드", type=["xlsx", "csv"])
target_file = uploaded_file if uploaded_file else (EXCEL_FILE if os.path.exists(EXCEL_FILE) else None)

def clean_series(series):
    cleaned = series.astype(str).str.replace(',', '').str.replace('₩', '').str.strip()
    cleaned = cleaned.replace(['-', '–', '—', 'nan', 'NaN', 'None', ''], '0')
    return pd.to_numeric(cleaned, errors='coerce').fillna(0)

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
# 4. '종합' 시트 정밀 파싱 (접수기준)
# =========================================================
summary_sheet_name = get_sheet_by_keyword(["종합"]) or excel_obj.sheet_names[0]
raw_summary = pd.read_excel(target_file, sheet_name=summary_sheet_name, header=None)

h_idx = 0
for idx, row in raw_summary.iterrows():
    r_text = "".join(row.dropna().astype(str).tolist())
    if "구분" in r_text and any(k in r_text for k in ["합계", "25", "26"]):
        h_idx = idx
        break

df_summary = pd.read_excel(target_file, sheet_name=summary_sheet_name, skiprows=h_idx)

for c in df_summary.columns:
    if df_summary[c].dtype == object:
        conv = pd.to_numeric(df_summary[c].astype(str).str.replace(',', '').str.strip(), errors='coerce')
        if conv.notnull().mean() > 0.5:
            df_summary[c] = conv.fillna(0)

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
# 5. 세부 파트 시트 우측 피벗 블록 정밀 파싱
# =========================================================
PART_SHEET_MAPPINGS = {
    "패션잡화": [["kc"]],
    "GB": [["gb"]],
    "글로벌 바이어": [["global", "1"], ["global", "2"]],
    "제품평가": [["inspection", "원단"], ["inspection", "가먼트"]]
}

def extract_pivot_block(sheet_name):
    raw = pd.read_excel(target_file, sheet_name=sheet_name, header=None)
    pivot_r, pivot_c = None, None
    for r_i in range(min(20, len(raw))):
        for c_i in range(len(raw.columns)):
            val = str(raw.iat[r_i, c_i]).strip().replace(" ", "")
            if "행레이블" in val:
                pivot_r, pivot_c = r_i, c_i
                break
        if pivot_r is not None:
            break
            
    if pivot_r is None:
        for r_i in range(min(20, len(raw))):
            row_str = " ".join(raw.iloc[r_i].dropna().astype(str).tolist())
            if "합계" in row_str and any(y in row_str for y in ["25", "26"]):
                pivot_r = r_i
                pivot_c = 0
                break
                
    if pivot_r is None:
        return pd.DataFrame(columns=["바이어명", "2025년 실적", "2026년 실적"])

    sub_raw = raw.iloc[pivot_r:, pivot_c:pivot_c+6].copy().reset_index(drop=True)
    sub_raw.columns = [str(c).strip() for c in sub_raw.iloc[0]]
    sub_data = sub_raw.iloc[1:].copy().reset_index(drop=True)
    
    buyer_col_name = sub_data.columns[0]
    c25, c26 = None, None
    for col in sub_data.columns[1:]:
        c_str = str(col).replace(" ", "")
        if "25" in c_str and ("합계" in c_str or "실적" in c_str):
            c25 = col
        elif "26" in c_str and ("합계" in c_str or "실적" in c_str):
            c26 = col
            
    if not c25 or not c26:
        num_candidates = [c for c in sub_data.columns[1:] if clean_series(sub_data[c]).sum() > 0]
        if len(num_candidates) >= 2:
            c25, c26 = num_candidates[0], num_candidates[1]
        else:
            return pd.DataFrame(columns=["바이어명", "2025년 실적", "2026년 실적"])

    parsed_rows = []
    for _, row in sub_data.iterrows():
        b_name = str(row[buyer_col_name]).strip()
        if not b_name or b_name.lower() in ['nan', 'none']:
            continue
        if any(k in b_name.replace(" ", "") for k in ["총합계", "합계", "전체합계"]):
            break
            
        v25 = clean_series(pd.Series([row[c25]])).iloc[0]
        v26 = clean_series(pd.Series([row[c26]])).iloc[0]
        parsed_rows.append({"바이어명": b_name, "2025년 실적": v25, "2026년 실적": v26})
        
    return pd.DataFrame(parsed_rows)

def get_combined_part_data(categories_target):
    keywords_list = PART_SHEET_MAPPINGS.get(categories_target, [])
    dfs = []
    for k_words in keywords_list:
        sheet_n = get_sheet_by_keyword(k_words)
        if sheet_n:
            block_df = extract_pivot_block(sheet_n)
            if not block_df.empty:
                dfs.append(block_df)
    if dfs:
        comb = pd.concat(dfs, ignore_index=True)
        return comb.groupby("바이어명", as_index=False)[["2025년 실적", "2026년 실적"]].sum()
    return pd.DataFrame(columns=["바이어명", "2025년 실적", "2026년 실적"])

part_data_cache = {cat: get_combined_part_data(cat) for cat in target_categories}

# =========================================================
# 6. BI_종합 시트 전용 파서: [월계] / [누계] 정밀 추출
# =========================================================
@st.cache_data
def get_bi_summary_kpi_and_chart(target_file_path):
    sheet_name = None
    for s_clean, orig_name in sheet_dict.items():
        if "bi" in s_clean and "종합" in s_clean:
            sheet_name = orig_name
            break
    if not sheet_name:
        for s_clean, orig_name in sheet_dict.items():
            if "bi" in s_clean:
                sheet_name = orig_name
                break

    if not sheet_name:
        # BI 시트 부재 시 접수기준 종합 실적으로 fallback
        t25 = float(summary_chart[col_25].sum())
        t26 = float(summary_chart[col_26].sum())
        return {
            "월계": {"25": t25, "26": t26, "chart": summary_chart.rename(columns={col_25: "25", col_26: "26"})},
            "누계": {"25": t25, "26": t26, "chart": summary_chart.rename(columns={col_25: "25", col_26: "26"})}
        }

    raw = pd.read_excel(target_file_path, sheet_name=sheet_name, header=None)

    # 1. 헤더 행 및 '월계' / '누계' 영역 탐색
    h_idx = 0
    for idx in range(min(25, len(raw))):
        row_str = " ".join(raw.iloc[idx].dropna().astype(str).tolist())
        if any(k in row_str for k in ["구분", "사업", "품목"]) and any(y in row_str for y in ["25", "26", "당월", "누계", "월계"]):
            h_idx = idx
            break

    df_bi = pd.read_excel(target_file_path, sheet_name=sheet_name, skiprows=h_idx)
    
    # 2. 총계 행 탐색
    total_rows = df_bi[df_bi.iloc[:, 0].astype(str).str.contains(r"총계|합계|TOTAL", regex=True, na=False)]
    
    # 컬럼 분류: '월계(당월)'와 '누계'
    month_cols_25, month_cols_26 = [], []
    cumul_cols_25, cumul_cols_26 = [], []

    for c in df_bi.columns:
        c_str = str(c).replace(" ", "")
        is_25 = "25" in c_str
        is_26 = "26" in c_str
        is_cumul = any(k in c_str for k in ["누계", "CUMUL"])
        
        if is_25:
            if is_cumul:
                cumul_cols_25.append(c)
            else:
                month_cols_25.append(c)
        elif is_26:
            if is_cumul:
                cumul_cols_26.append(c)
            else:
                month_cols_26.append(c)

    # 기본값 및 총계 계산
    def extract_totals_and_chart(c25_list, c26_list, fallback_25, fallback_26):
        c25 = c25_list[0] if c25_list else None
        c26 = c26_list[0] if c26_list else None
        
        val_25 = fallback_25
        val_26 = fallback_26
        
        if not total_rows.empty and c25 and c26:
            val_25 = clean_series(total_rows[c25]).iloc[0]
            val_26 = clean_series(total_rows[c26]).iloc[0]
            
        # 사업별 차트용 데이터
        chart_data = summary_chart.copy().rename(columns={col_25: "25", col_26: "26"})
        if c25 and c26:
            # 사업구분 매핑
            first_col = df_bi.columns[0]
            df_bi_clean = df_bi[~df_bi[first_col].astype(str).str.contains(r"총계|합계|TOTAL", regex=True, na=False)].copy()
            df_bi_clean["표준"] = df_bi_clean[first_col].apply(map_biz_category)
            filtered = df_bi_clean.dropna(subset=["표준"])
            if not filtered.empty:
                chart_data = filtered.groupby("표준", as_index=False)[[c25, c26]].sum().rename(columns={"표준": "표준사업구분", c25: "25", c26: "26"})

        return {"25": val_25, "26": val_26, "chart": chart_data}

    # 접수기준 종합합계
    def_25 = float(summary_chart[col_25].sum())
    def_26 = float(summary_chart[col_26].sum())

    month_pack = extract_totals_and_chart(month_cols_25, month_cols_26, def_25, def_26)
    cumul_pack = extract_totals_and_chart(cumul_cols_25, cumul_cols_26, def_25, def_26)

    return {"월계": month_pack, "누계": cumul_pack}

bi_data_pack = get_bi_summary_kpi_and_chart(target_file)

# =========================================================
# 7. 사이드바 메뉴: 분석 페이지 선택
# =========================================================
st.sidebar.markdown("### 📑 분석 페이지 선택")
page_menu = st.sidebar.radio(
    "이동할 카테고리를 선택하세요:",
    [
        "[접수기준] 종합 실적 현황",
        "[접수기준] 사업별 실적 현황",
        "[접수기준] 바이어 실적 현황",
        "[BI_종합] 사업별 실적 현황",
        "[BI_상해+광주] 사업별 실적 현황"
    ],
    index=0
)

# =========================================================
# 8. 상단 종합 KPI 카드 (BI_종합 월계/누계 토글 및 사업별 연동)
# =========================================================
# 기본 KPI 계산 대상
selected_view_for_card = "전체 사업 보기"
bi_period_mode = "누계"

# 상단 우측에 월계/누계 선택 토글 배치
st.sidebar.markdown("---")
st.sidebar.markdown("### ⏱️ 실적 집계 기준 (BI_종합)")
bi_period_mode = st.sidebar.radio("집계 구분 선택:", ["누계", "월계"], index=0)

if page_menu == "[접수기준] 사업별 실적 현황":
    if "selected_biz_view" not in st.session_state:
        st.session_state["selected_biz_view"] = "전체 사업 보기"
    selected_view_for_card = st.session_state["selected_biz_view"]
elif page_menu == "[접수기준] 바이어 실적 현황":
    if "selected_tab3_biz" not in st.session_state:
        st.session_state["selected_tab3_biz"] = target_categories[0]
    selected_view_for_card = st.session_state["selected_tab3_biz"]

# 카드 수치 결정
if page_menu.startswith("[BI_종합]"):
    cur_pack = bi_data_pack.get(bi_period_mode, bi_data_pack["누계"])
    total_25 = float(cur_pack["25"])
    total_26 = float(cur_pack["26"])
    card_sub_desc = f"BI_종합 [{bi_period_mode}] 총계 기준"
elif selected_view_for_card != "전체 사업 보기" and selected_view_for_card in target_categories:
    target_row = summary_chart[summary_chart["표준사업구분"] == selected_view_for_card]
    total_25 = float(target_row[col_25].sum()) if not target_row.empty else 0.0
    total_26 = float(target_row[col_26].sum()) if not target_row.empty else 0.0
    card_sub_desc = f"[{selected_view_for_card}] 실적 합계"
else:
    # 기본 전체보기
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
        <div class="kpi-title">📅 25년 총 실적 ({bi_period_mode if page_menu.startswith('[BI_') else '누적'})</div>
        <div class="kpi-num">{total_25:,.0f}</div>
        <div class="kpi-sub">{card_sub_desc}</div>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown(f"""
    <div class="kpi-card" style="border-top-color: #003876;">
        <div class="kpi-title">🚀 26년 총 실적 ({bi_period_mode if page_menu.startswith('[BI_') else '누적'})</div>
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
# 9. 본문 페이지 렌더링
# =========================================================

# [페이지 1] [접수기준] 종합 실적 현황
if page_menu == "[접수기준] 종합 실적 현황":
    st.subheader("📌 2025년 총 실적 vs 2026년 총 실적 비교 (접수기준)")
    
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

# [페이지 2] [접수기준] 사업별 실적 현황
elif page_menu == "[접수기준] 사업별 실적 현황":
    st.subheader("🏢 사업별 2025년 vs 2026년 실적 증감 비교 (접수기준)")
    
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

# [페이지 3] [접수기준] 바이어 실적 현황
elif page_menu == "[접수기준] 바이어 실적 현황":
    st.subheader("🤝 각 사업별 주요 바이어 2025년 vs 2026년 실적 변화 (접수기준)")
    
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
    
    target_row = summary_chart[summary_chart["표준사업구분"] == selected_biz]
    target_tot_25 = float(target_row[col_25].values[0]) if not target_row.empty else 0.0
    target_tot_26 = float(target_row[col_26].values[0]) if not target_row.empty else 0.0

    if raw_b_chart.empty:
        st.warning(f"선택하신 [{selected_biz}] 부문의 요약 블록('행 레이블' 표)을 읽을 수 없습니다.")
    else:
        is_dash = raw_b_chart["바이어명"].astype(str).str.strip().isin(["-", "–", "—", "", "NAN", "NONE", "기타"])
        valid_buyers = raw_b_chart[~is_dash].copy()
        dash_buyers = raw_b_chart[is_dash].copy()
        
        valid_buyers = valid_buyers.sort_values(by="2026년 실적", ascending=False).reset_index(drop=True)
        
        if len(valid_buyers) > 6:
            top6 = valid_buyers.iloc[:6].copy()
            rest = valid_buyers.iloc[6:].copy()
            
            top6_25 = top6["2025년 실적"].sum()
            top6_26 = top6["2026년 실적"].sum()
            
            calc_etc_25 = rest["2025년 실적"].sum() + dash_buyers["2025년 실적"].sum()
            calc_etc_26 = rest["2026년 실적"].sum() + dash_buyers["2026년 실적"].sum()
            
            etc_25 = max(calc_etc_25, target_tot_25 - top6_25)
            etc_26 = max(calc_etc_26, target_tot_26 - top6_26)
            
            etc_row = pd.DataFrame([{
                "바이어명": "기타",
                "2025년 실적": etc_25,
                "2026년 실적": etc_26
            }])
            top_buyers = pd.concat([top6, etc_row], ignore_index=True)
        else:
            top_sum_25 = valid_buyers["2025년 실적"].sum()
            top_sum_26 = valid_buyers["2026년 실적"].sum()
            
            etc_25 = max(dash_buyers["2025년 실적"].sum(), target_tot_25 - top_sum_25)
            etc_26 = max(dash_buyers["2026년 실적"].sum(), target_tot_26 - top_sum_26)
            
            if etc_25 > 0 or etc_26 > 0:
                etc_row = pd.DataFrame([{
                    "바이어명": "기타",
                    "2025년 실적": etc_25,
                    "2026년 실적": etc_26
                }])
                top_buyers = pd.concat([valid_buyers, etc_row], ignore_index=True)
            else:
                top_buyers = valid_buyers.copy()

        top_buyers["증감액"] = top_buyers["2026년 실적"] - top_buyers["2025년 실적"]
        top_buyers["증감률"] = ((top_buyers["증감액"] / top_buyers["2025년 실적"].replace(0, pd.NA)) * 100).fillna(0.0)

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

# [페이지 4] [BI_종합] 사업별 실적 현황
elif page_menu == "[BI_종합] 사업별 실적 현황":
    st.subheader(f"📊 [BI_종합] 사업별 2025년 vs 2026년 실적 비교 ({bi_period_mode} 기준)")
    
    cur_pack = bi_data_pack.get(bi_period_mode, bi_data_pack["누계"])
    bi_df = cur_pack["chart"].copy()
    
    bi_df["증감액"] = bi_df["26"] - bi_df["25"]
    bi_df["증감률"] = ((bi_df["증감액"] / bi_df["25"].replace(0, pd.NA)) * 100).fillna(0.0)
    
    col4_l, col4_r = st.columns([6, 4])
    
    with col4_l:
        fig4_bar = go.Figure()
        fig4_bar.add_trace(go.Bar(
            x=bi_df["표준사업구분"],
            y=bi_df["25"],
            name="2025년 실적",
            marker=dict(color="#94A3B8", line=dict(color="#64748B", width=1), cornerradius=6),
            text=bi_df["25"].apply(lambda x: f"{x/1e8:.1f}억" if x >= 1e8 else f"{x/1e4:.0f}만"),
            textposition="outside",
            textfont=dict(size=11, color="#475569", family="Pretendard")
        ))
        fig4_bar.add_trace(go.Bar(
            x=bi_df["표준사업구분"],
            y=bi_df["26"],
            name="2026년 실적",
            marker=dict(color="#1D4ED8", line=dict(color="#1E40AF", width=1), cornerradius=6),
            text=bi_df["26"].apply(lambda x: f"{x/1e8:.1f}억" if x >= 1e8 else f"{x/1e4:.0f}만"),
            textposition="outside",
            textfont=dict(size=11, color="#0F172A", family="Pretendard", weight="bold")
        ))
        fig4_bar.update_layout(
            height=440,
            bargap=0.30,
            bargroupgap=0.10,
            yaxis=dict(rangemode='tozero', title=dict(text="실적금액 (원)", font=dict(size=12, color="#64748B")), gridcolor="#F1F5F9"),
            xaxis=dict(categoryorder='array', categoryarray=target_categories, tickfont=dict(size=13, weight="bold", color="#1E293B")),
            template="plotly_white",
            legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="left", x=0),
            margin=dict(t=50, b=20, l=10, r=10)
        )
        st.plotly_chart(fig4_bar, use_container_width=True)

    with col4_r:
        diff_colors = ["#E11D48" if v >= 0 else "#2563EB" for v in bi_df["증감액"]]
        diff_texts = [f"{'+' if v >= 0 else ''}{v/1e8:.2f}억" if abs(v) >= 1e8 else f"{'+' if v >= 0 else ''}{v/1e4:.0f}만" for v in bi_df["증감액"]]
        
        fig4_diff = go.Figure()
        fig4_diff.add_trace(go.Bar(
            x=bi_df["표준사업구분"],
            y=bi_df["증감액"],
            marker=dict(color=diff_colors, cornerradius=6),
            text=diff_texts,
            textposition="outside",
            textfont=dict(size=11, weight="bold", family="Pretendard")
        ))
        fig4_diff.update_layout(
            title=f"[BI_종합 - {bi_period_mode}] 사업별 증감액 (26년 - 25년)",
            height=440,
            bargap=0.45,
            yaxis=dict(title=dict(text="증감액 (원)", font=dict(size=12, color="#64748B")), gridcolor="#F1F5F9", zerolinecolor="#CBD5E1"),
            xaxis=dict(categoryorder='array', categoryarray=target_categories, tickfont=dict(size=13, weight="bold", color="#1E293B")),
            template="plotly_white",
            margin=dict(t=50, b=20, l=10, r=10)
        )
        st.plotly_chart(fig4_diff, use_container_width=True)
        
    st.markdown(f"##### 📋 [BI_종합 - {bi_period_mode}] 사업별 실적 요약 테이블")
    st.dataframe(
        pd.DataFrame({
            "사업구분": bi_df["표준사업구분"],
            "2025년 실적 (원)": bi_df["25"],
            "2026년 실적 (원)": bi_df["26"],
            "증감액 (원)": bi_df["증감액"],
            "증감률(%)": bi_df["증감률"]
        }),
        column_config={
            "사업구분": st.column_config.TextColumn("사업구분", width="medium"),
            "2025년 실적 (원)": st.column_config.NumberColumn("2025년 실적 (원)", format="₩%,d"),
            "2026년 실적 (원)": st.column_config.NumberColumn("2026년 실적 (원)", format="₩%,d"),
            "증감액 (원)": st.column_config.NumberColumn("증감액 (원)", format="₩%+d"),
            "증감률(%)": st.column_config.NumberColumn("증감률(%)", format="%+.2f%%"),
        },
        hide_index=True,
        use_container_width=True
    )

# [페이지 5] [BI_상해+광주] 사업별 실적 현황
elif page_menu == "[BI_상해+광주] 사업별 실적 현황":
    st.subheader("🌐 [BI_상해+광주] 통합 사업별 2025년 vs 2026년 실적 비교")
    
    bi_gz_chart = summary_chart.copy()
    gz_c25, gz_c26 = col_25, col_26
    
    bi_gz_chart["증감액"] = bi_gz_chart[gz_c26] - bi_gz_chart[gz_c25]
    bi_gz_chart["증감률"] = ((bi_gz_chart["증감액"] / bi_gz_chart[gz_c25].replace(0, pd.NA)) * 100).fillna(0.0)
    
    col5_l, col5_r = st.columns([6, 4])
    
    with col5_l:
        fig5_bar = go.Figure()
        fig5_bar.add_trace(go.Bar(
            x=bi_gz_chart["표준사업구분"],
            y=bi_gz_chart[gz_c25],
            name="2025년 실적",
            marker=dict(color="#94A3B8", line=dict(color="#64748B", width=1), cornerradius=6),
            text=bi_gz_chart[gz_c25].apply(lambda x: f"{x/1e8:.1f}억" if x >= 1e8 else f"{x/1e4:.0f}만"),
            textposition="outside",
            textfont=dict(size=11, color="#475569", family="Pretendard")
        ))
        fig5_bar.add_trace(go.Bar(
            x=bi_gz_chart["표준사업구분"],
            y=bi_gz_chart[gz_c26],
            name="2026년 실적",
            marker=dict(color="#1D4ED8", line=dict(color="#1E40AF", width=1), cornerradius=6),
            text=bi_gz_chart[gz_c26].apply(lambda x: f"{x/1e8:.1f}억" if x >= 1e8 else f"{x/1e4:.0f}만"),
            textposition="outside",
            textfont=dict(size=11, color="#0F172A", family="Pretendard", weight="bold")
        ))
        fig5_bar.update_layout(
            height=440,
            bargap=0.30,
            bargroupgap=0.10,
            yaxis=dict(rangemode='tozero', title=dict(text="실적금액 (원)", font=dict(size=12, color="#64748B")), gridcolor="#F1F5F9"),
            xaxis=dict(categoryorder='array', categoryarray=target_categories, tickfont=dict(size=13, weight="bold", color="#1E293B")),
            template="plotly_white",
            legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="left", x=0),
            margin=dict(t=50, b=20, l=10, r=10)
        )
        st.plotly_chart(fig5_bar, use_container_width=True)

    with col5_r:
        diff_colors = ["#E11D48" if v >= 0 else "#2563EB" for v in bi_gz_chart["증감액"]]
        diff_texts = [f"{'+' if v >= 0 else ''}{v/1e8:.2f}억" if abs(v) >= 1e8 else f"{'+' if v >= 0 else ''}{v/1e4:.0f}만" for v in bi_gz_chart["증감액"]]
        
        fig5_diff = go.Figure()
        fig5_diff.add_trace(go.Bar(
            x=bi_gz_chart["표준사업구분"],
            y=bi_gz_chart["증감액"],
            marker=dict(color=diff_colors, cornerradius=6),
            text=diff_texts,
            textposition="outside",
            textfont=dict(size=11, weight="bold", family="Pretendard")
        ))
        fig5_diff.update_layout(
            title="[BI_상해+광주] 사업별 증감액 (26년 - 25년)",
            height=440,
            bargap=0.45,
            yaxis=dict(title=dict(text="증감액 (원)", font=dict(size=12, color="#64748B")), gridcolor="#F1F5F9", zerolinecolor="#CBD5E1"),
            xaxis=dict(categoryorder='array', categoryarray=target_categories, tickfont=dict(size=13, weight="bold", color="#1E293B")),
            template="plotly_white",
            margin=dict(t=50, b=20, l=10, r=10)
        )
        st.plotly_chart(fig5_diff, use_container_width=True)
        
    st.markdown("##### 📋 [BI_상해+광주] 사업별 실적 요약 테이블")
    st.dataframe(
        pd.DataFrame({
            "사업구분": bi_gz_chart["표준사업구분"],
            "2025년 실적 (원)": bi_gz_chart[gz_c25],
            "2026년 실적 (원)": bi_gz_chart[gz_c26],
            "증감액 (원)": bi_gz_chart["증감액"],
            "증감률(%)": bi_gz_chart["증감률"]
        }),
        column_config={
            "사업구분": st.column_config.TextColumn("사업구분", width="medium"),
            "2025년 실적 (원)": st.column_config.NumberColumn("2025년 실적 (원)", format="₩%,d"),
            "2026년 실적 (원)": st.column_config.NumberColumn("2026년 실적 (원)", format="₩%,d"),
            "증감액 (원)": st.column_config.NumberColumn("증감액 (원)", format="₩%+d"),
            "증감률(%)": st.column_config.NumberColumn("증감률(%)", format="%+.2f%%"),
        },
        hide_index=True,
        use_container_width=True
    )
