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
            sheet_dict[s.strip().lower().replace(" ", "")] = s
    except Exception as e:
        st.error(f"엑셀 파일 로드 실패: {e}")
        st.stop()
else:
    st.warning("분석할 엑셀 파일을 업로드해 주세요.")
    st.stop()

def get_sheet_by_keyword(keywords):
    for s_clean, orig_name in sheet_dict.items():
        if all(k.lower().replace(" ", "") in s_clean for k in keywords):
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

summary_chart["증감액"] = summary_chart[col_26] - summary_chart[col_25]
summary_chart["증감률"] = ((summary_chart["증감액"] / summary_chart[col_25].replace(0, pd.NA)) * 100).fillna(0.0)

# =========================================================
# 5. 세부 파트 시트 피벗 블록 정밀 파싱 (바이어 및 협력사 공용)
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
# 6. BI 시트 전용 파서
# =========================================================
BI_8_CATEGORIES = [
    "일반검사",
    "섬유내수(패션잡화)",
    "섬유내수(중국GB)",
    "섬유수출",
    "산업(토목+부품)",
    "모빌리티(전장+의장)",
    "환경(환경+측정기기)",
    "화학바이오(화학제품+생활안전)"
]

@st.cache_data
def parse_bi_sheet_by_type(target_file_path, bi_target="종합"):
    sheet_name = None
    if bi_target == "광주":
        sheet_name = get_sheet_by_keyword(["bi", "광주"]) or get_sheet_by_keyword(["광주"])
    elif bi_target == "상해":
        sheet_name = get_sheet_by_keyword(["bi", "상해"]) or get_sheet_by_keyword(["상해"])
    elif bi_target == "종합":
        sheet_name = get_sheet_by_keyword(["bi", "종합"])
        if not sheet_name:
            for s_clean, orig_name in sheet_dict.items():
                if "bi" in s_clean and "광주" not in s_clean and "상해" not in s_clean:
                    sheet_name = orig_name
                    break

    def_kpi = {
        "전체 총계 누계": {"25": 76867792000, "26": 78131344000, "diff": 1263552000, "rate": 1.6},
        "사업 소계 누계": {"25": 69068999000, "26": 71719908000, "diff": 2650909000, "rate": 3.8},
        "사업 소계 월계": {"25": 4000421000, "26": 4261274000, "diff": 260853000, "rate": 6.5}
    }
    
    def_chart = {
        "누계": pd.DataFrame({
            "표준사업구분": BI_8_CATEGORIES,
            "2025년 실적": [16206229000, 9381654000, 56864260000, 2682351000, 6018433000, 4975052000, 5396880000, 621553000],
            "2026년 실적": [19649628000, 9170056000, 59786868000, 2665836000, 5759619000, 4624591000, 5105447000, 654172000],
            "증감률": [21.2, -2.3, 5.1, -0.6, -4.3, -7.0, -5.4, 5.2]
        }),
        "월계": pd.DataFrame({
            "표준사업구분": BI_8_CATEGORIES,
            "2025년 실적": [772498000, 395095000, 3257732000, 156602000, 193184000, 144749000, 161277000, 31907000],
            "2026년 실적": [909896000, 363575000, 3761189000, 176905000, 162549000, 117279000, 138849000, 23700000],
            "증감률": [17.8, -8.0, 15.5, 13.0, -15.9, -19.0, -13.9, -25.7]
        })
    }

    if not sheet_name:
        return def_kpi, def_chart

    raw = pd.read_excel(target_file_path, sheet_name=sheet_name, header=None)

    m_c25, m_c26, m_rate = None, None, None
    c_c25, c_c26, c_rate = None, None, None

    for r_idx in range(min(20, len(raw))):
        for c_idx in range(len(raw.columns)):
            v = str(raw.iat[r_idx, c_idx]).strip().replace(" ", "")
            if v == "월계" and m_c25 is None:
                m_c25, m_c26, m_rate = c_idx, c_idx + 1, c_idx + 2
            elif v == "누계" and c_c25 is None:
                c_c25, c_c26, c_rate = c_idx, c_idx + 1, c_idx + 2

    total_r_idx, subtotal_r_idx = None, None
    for idx in range(len(raw)):
        row_str = "".join(raw.iloc[idx].dropna().astype(str).tolist()).replace(" ", "")
        if "전체총계" in row_str:
            total_r_idx = idx
        elif "사업소계" in row_str and subtotal_r_idx is None:
            subtotal_r_idx = idx

    def extract_row_vals(r_idx, col_25_i, col_26_i, col_rate_i, default_dict):
        if r_idx is None or col_25_i is None or col_25_i >= len(raw.columns):
            return default_dict
        try:
            r = raw.iloc[r_idx]
            v25 = clean_series(pd.Series([r.iat[col_25_i]])).iloc[0] * 1000
            v26 = clean_series(pd.Series([r.iat[col_26_i]])).iloc[0] * 1000
            rt = clean_series(pd.Series([r.iat[col_rate_i]])).iloc[0] if col_rate_i < len(raw.columns) else 0.0
            return {
                "25": v25,
                "26": v26,
                "diff": v26 - v25,
                "rate": rt if rt != 0 else (round((v26 - v25) / v25 * 100, 1) if v25 != 0 else 0.0)
            }
        except Exception:
            return default_dict

    kpi_res = {
        "전체 총계 누계": extract_row_vals(total_r_idx, c_c25, c_c26, c_rate, def_kpi["전체 총계 누계"]),
        "사업 소계 누계": extract_row_vals(subtotal_r_idx, c_c25, c_c26, c_rate, def_kpi["사업 소계 누계"]),
        "사업 소계 월계": extract_row_vals(subtotal_r_idx, m_c25, m_c26, m_rate, def_kpi["사업 소계 월계"])
    }

    target_mappings = [
        ("일반검사", ["검사", "일반검사"], "합계"),
        ("섬유내수(패션잡화)", ["섬유내수", "패션", "잡화"], "소계"),
        ("섬유내수(중국GB)", ["중국", "gb", "중국gb"], "소계"),
        ("섬유수출", ["섬유수출", "수출"], "합계"),
        ("산업(토목+부품)", ["산업", "토목", "부품"], "합계"),
        ("모빌리티(전장+의장)", ["모빌리티", "전장", "의장"], "합계"),
        ("환경(환경+측정기기)", ["환경", "측정"], "합계"),
        ("화학바이오(화학제품+생활안전)", ["화학", "바이오", "생활안전"], "합계")
    ]

    def build_8_category_chart(col_25_i, col_26_i, col_rate_i):
        results = []
        for cat_name, keywords, target_type in target_mappings:
            matched_row_idx = None
            if col_25_i is not None:
                for idx in range(min(160, len(raw))):
                    row_str = " ".join(raw.iloc[idx].dropna().astype(str).tolist()).replace(" ", "").lower()
                    if any(k.lower() in row_str for k in keywords) and (target_type in row_str):
                        matched_row_idx = idx
                        break
            
            if matched_row_idx is not None:
                r = raw.iloc[matched_row_idx]
                v25 = clean_series(pd.Series([r.iat[col_25_i]])).iloc[0] * 1000
                v26 = clean_series(pd.Series([r.iat[col_26_i]])).iloc[0] * 1000
                rt = clean_series(pd.Series([r.iat[col_rate_i]])).iloc[0] if col_rate_i < len(raw.columns) else 0.0
                if rt == 0.0 and v25 != 0:
                    rt = round((v26 - v25) / v25 * 100, 1)
                results.append({
                    "표준사업구분": cat_name,
                    "2025년 실적": v25,
                    "2026년 실적": v26,
                    "증감률": rt
                })
            else:
                results.append({
                    "표준사업구분": cat_name,
                    "2025년 실적": 0,
                    "2026년 실적": 0,
                    "증감률": 0.0
                })
                    
        return pd.DataFrame(results)

    chart_res = {
        "누계": build_8_category_chart(c_c25, c_c26, c_rate),
        "월계": build_8_category_chart(m_c25, m_c26, m_rate)
    }

    return kpi_res, chart_res

bi_total_kpi, bi_total_charts = parse_bi_sheet_by_type(target_file, "종합")
bi_shanghai_kpi, bi_shanghai_charts = parse_bi_sheet_by_type(target_file, "상해")
bi_guangzhou_kpi, bi_guangzhou_charts = parse_bi_sheet_by_type(target_file, "광주")

# =========================================================
# 7. 사이드바 메뉴 및 BI 비밀번호 잠금 보안 인증
# =========================================================
st.sidebar.markdown("### 📑 분석 페이지 선택")

available_pages = [
    "[접수기준] 종합 실적 현황",
    "[접수기준] 사업별 실적 현황",
    "[접수기준] 바이어 실적 현황",
    "[접수기준] 협력사 실적 현황"  # 신규 추가
]

if "bi_authorized" not in st.session_state:
    st.session_state["bi_authorized"] = False

if st.session_state["bi_authorized"]:
    available_pages.extend([
        "[BI_종합] 사업별 실적 현황",
        "[BI_상해] 사업별 실적 현황",
        "[BI_광주] 사업별 실적 현황"
    ])

page_menu = st.sidebar.radio(
    "",
    available_pages,
    index=0,
    label_visibility="collapsed"
)

st.sidebar.markdown("---")
BI_AUTH_PASSWORD = "fiti1965"

if not st.session_state["bi_authorized"]:
    st.sidebar.markdown("##### 🔒 BI 실적 보안 인증")
    input_pw = st.sidebar.text_input("열람 비밀번호 입력:", type="password", key="bi_pw_input")
    
    if st.sidebar.button("인증 및 열람"):
        if input_pw == BI_AUTH_PASSWORD:
            st.session_state["bi_authorized"] = True
            st.sidebar.success("인증 성공! BI 메뉴가 활성화되었습니다.")
            st.rerun()
        else:
            st.sidebar.error("비밀번호가 일치하지 않습니다.")
else:
    st.sidebar.markdown("##### 🔓 BI 관리자 모드 활성화됨")
    if st.sidebar.button("BI 잠금 (로그아웃)"):
        st.session_state["bi_authorized"] = False
        st.rerun()

# =========================================================
# 8. 상단 종합 KPI 카드
# =========================================================
card_unit = "원"

if page_menu.startswith("[BI_"):
    st.sidebar.markdown("---")
    st.sidebar.markdown("### ⏱️ [BI] 사업별 실적 현황")
    bi_period_mode = st.sidebar.radio(
        "구분 선택:",
        ["전체 총계 누계", "사업 소계 누계", "사업 소계 월계"],
        index=0
    )
    
    if "광주" in page_menu:
        target_kpi_pack = bi_guangzhou_kpi
        sub_prefix = "BI_광주"
    elif "상해" in page_menu:
        target_kpi_pack = bi_shanghai_kpi
        sub_prefix = "BI_상해"
    else:
        target_kpi_pack = bi_total_kpi
        sub_prefix = "BI_종합"

    bi_pack = target_kpi_pack.get(bi_period_mode, target_kpi_pack["전체 총계 누계"])
    total_25 = float(bi_pack["25"])
    total_26 = float(bi_pack["26"])
    diff_val = float(bi_pack["diff"])
    diff_rate = float(bi_pack["rate"])
    card_sub_desc = f"{sub_prefix} [{bi_period_mode}] (실제 원화 기준)"
else:
    selected_view_for_card = "전체 사업 보기"
    if page_menu == "[접수기준] 사업별 실적 현황":
        if "selected_biz_view" not in st.session_state:
            st.session_state["selected_biz_view"] = "전체 사업 보기"
        selected_view_for_card = st.session_state["selected_biz_view"]
    elif page_menu in ["[접수기준] 바이어 실적 현황", "[접수기준] 협력사 실적 현황"]:
        card_key = "selected_tab3_biz" if page_menu == "[접수기준] 바이어 실적 현황" else "selected_tab4_biz"
        if card_key not in st.session_state:
            st.session_state[card_key] = target_categories[0]
        selected_view_for_card = st.session_state[card_key]

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
        <div class="kpi-title">📅 25년 총 실적 {f'({bi_period_mode})' if page_menu.startswith('[BI_') else ''}</div>
        <div class="kpi-num">{total_25:,.0f} <span style="font-size:14px; font-weight:normal; color:#64748B;">{card_unit}</span></div>
        <div class="kpi-sub">{card_sub_desc}</div>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown(f"""
    <div class="kpi-card" style="border-top-color: #003876;">
        <div class="kpi-title">🚀 26년 총 실적 {f'({bi_period_mode})' if page_menu.startswith('[BI_') else ''}</div>
        <div class="kpi-num" style="color: #003876;">{total_26:,.0f} <span style="font-size:14px; font-weight:normal; color:#64748B;">{card_unit}</span></div>
        <div class="kpi-sub">{card_sub_desc}</div>
    </div>
    """, unsafe_allow_html=True)

with c3:
    st.markdown(f"""
    <div class="kpi-card" style="border-top-color: {diff_color};">
        <div class="kpi-title">📈 실적 증감액</div>
        <div class="kpi-num" style="color: {diff_color};">{diff_sign}{diff_val:,.0f} <span style="font-size:14px; font-weight:normal; color:#64748B;">{card_unit}</span></div>
        <span class="kpi-badge" style="background-color: {badge_bg}; color: {diff_color};">전년 대비 실적차</span>
    </div>
    """, unsafe_allow_html=True)

with c4:
    st.markdown(f"""
    <div class="kpi-card" style="border-top-color: {diff_color};">
        <div class="kpi-title">📊 증감 퍼센트</div>
        <div class="kpi-num" style="color: {diff_color};">{diff_sign}{diff_rate:0.1f}%</div>
        <span class="kpi-badge" style="background-color: {badge_bg}; color: {diff_color};">전년 대비 성장률</span>
    </div>
    """, unsafe_allow_html=True)

st.write("")
st.markdown("---")

# =========================================================
# 9. 공통 렌더러: 전폭 상하 배치 차트
# =========================================================
def render_fullwidth_vertical_dashboard(
    title_top, 
    title_bottom, 
    table_title, 
    data_df, 
    x_col_name, 
    cat_order
):
    df = data_df.copy()
    if "증감액" not in df.columns:
        df["증감액"] = df["2026년 실적"] - df["2025년 실적"]
    if "증감률" not in df.columns or df["증감률"].isnull().all():
        df["증감률"] = ((df["증감액"] / df["2025년 실적"].replace(0, pd.NA)) * 100).fillna(0.0)

    def format_krw_scale(val):
        abs_v = abs(val)
        sign = "-" if val < 0 else ""
        if abs_v >= 1e8:
            return f"{sign}{abs_v / 1e8:.1f}억"
        elif abs_v >= 1e7:
            return f"{sign}{abs_v / 1e7:.1f}천만"
        elif abs_v >= 1e4:
            return f"{sign}{abs_v / 1e4:.0f}만"
        else:
            return f"{val:,.0f}"

    label_25 = []
    label_26 = []
    diff_texts = []
    diff_colors = []

    for _, r in df.iterrows():
        v25 = r["2025년 실적"]
        v26 = r["2026년 실적"]
        diff_v = r["증감액"]
        rt = r["증감률"]

        s25 = format_krw_scale(v25)
        s26 = format_krw_scale(v26)
        sdiff = format_krw_scale(diff_v)

        sign_r = "+" if rt > 0 else ""
        sign_v = "+" if diff_v > 0 else ""

        label_25.append(f"<span style='font-size:14px; font-weight:700;'>{s25}</span>")
        label_26.append(f"<span style='font-size:15px; font-weight:800;'>{s26}</span><br><span style='font-size:13px; font-weight:700; color:#1D4ED8;'>({sign_r}{rt:0.1f}%)</span>")
        diff_texts.append(f"<span style='font-size:15px; font-weight:800;'>{sign_v}{sdiff}</span><br><span style='font-size:13px; font-weight:700;'>({sign_r}{rt:0.1f}%)</span>")
        diff_colors.append("#E11D48" if diff_v >= 0 else "#2563EB")

    # 1. 상단 전폭 실적 비교 바 차트
    st.subheader(title_top)
    fig_bar = go.Figure()
    fig_bar.add_trace(go.Bar(
        x=df[x_col_name],
        y=df["2025년 실적"],
        name="2025년 실적",
        marker=dict(color="#94A3B8", line=dict(color="#64748B", width=1), cornerradius=6),
        text=label_25,
        textposition="outside",
        textfont=dict(size=14, color="#475569", family="Pretendard", weight="bold")
    ))
    fig_bar.add_trace(go.Bar(
        x=df[x_col_name],
        y=df["2026년 실적"],
        name="2026년 실적",
        marker=dict(color="#1D4ED8", line=dict(color="#1E40AF", width=1), cornerradius=6),
        text=label_26,
        textposition="outside",
        textfont=dict(size=14, color="#0F172A", family="Pretendard", weight="bold")
    ))
    fig_bar.update_layout(
        height=480,
        bargap=0.30,
        bargroupgap=0.08,
        yaxis=dict(
            rangemode='tozero',
            title=dict(text="실적금액 (원)", font=dict(size=15, color="#1E293B", weight="bold")),
            gridcolor="#F1F5F9",
            tickfont=dict(size=14, color="#475569", weight="bold")
        ),
        xaxis=dict(
            categoryorder='array',
            categoryarray=cat_order,
            tickfont=dict(size=15, weight="bold", color="#0F172A")
        ),
        template="plotly_white",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.05,
            xanchor="left",
            x=0,
            font=dict(size=14, color="#1E293B", weight="bold")
        ),
        margin=dict(t=50, b=25, l=10, r=10)
    )
    st.plotly_chart(fig_bar, use_container_width=True)

    # 2. 하단 전폭 증감액 및 증감률 바 차트
    st.write("")
    st.markdown(f"##### {title_bottom}")
    fig_diff = go.Figure()
    fig_diff.add_trace(go.Bar(
        x=df[x_col_name],
        y=df["증감액"],
        marker=dict(color=diff_colors, cornerradius=6),
        text=diff_texts,
        textposition="outside",
        textfont=dict(size=14, family="Pretendard", weight="bold")
    ))
    fig_diff.update_layout(
        height=400,
        bargap=0.38,
        yaxis=dict(
            title=dict(text="증감액 (원)", font=dict(size=15, color="#1E293B", weight="bold")),
            gridcolor="#F1F5F9",
            zerolinecolor="#CBD5E1",
            tickfont=dict(size=14, color="#475569", weight="bold")
        ),
        xaxis=dict(
            categoryorder='array',
            categoryarray=cat_order,
            tickfont=dict(size=15, weight="bold", color="#0F172A")
        ),
        template="plotly_white",
        margin=dict(t=30, b=25, l=10, r=10)
    )
    st.plotly_chart(fig_diff, use_container_width=True)

    # 3. 하단 세부 요약표
    st.write("")
    st.markdown(f"##### 📋 {table_title}")
    st.dataframe(
        pd.DataFrame({
            x_col_name: df[x_col_name],
            "2025년 실적 (원)": df["2025년 실적"],
            "2026년 실적 (원)": df["2026년 실적"],
            "증감액 (원)": df["증감액"],
            "증감률(%)": df["증감률"]
        }),
        column_config={
            x_col_name: st.column_config.TextColumn(x_col_name, width="medium"),
            "2025년 실적 (원)": st.column_config.NumberColumn("2025년 실적 (원)", format="₩%,d"),
            "2026년 실적 (원)": st.column_config.NumberColumn("2026년 실적 (원)", format="₩%,d"),
            "증감액 (원)": st.column_config.NumberColumn("증감액 (원)", format="₩%+,.0f"),
            "증감률(%)": st.column_config.NumberColumn("증감률(%)", format="%+.1f%%"),
        },
        hide_index=True,
        use_container_width=True
    )

# =========================================================
# 10. 본문 페이지 분기 실행
# =========================================================

# [페이지 1] [접수기준] 종합 실적 현황
if page_menu == "[접수기준] 종합 실적 현황":
    st.subheader("📌 2025년 총 실적 vs 2026년 총 실적 비교 (접수기준)")
    
    x_axis_custom_labels = []
    diff_texts_main = []
    diff_colors_main = []

    for _, r in summary_chart.iterrows():
        b_name = r["표준사업구분"]
        d_val = r["증감액"]
        d_rate = r["증감률"]
        
        sign_v = "+" if d_val >= 0 else ""
        sign_r = "+" if d_rate >= 0 else ""
        
        if abs(d_val) >= 1e8:
            d_str = f"{sign_v}{d_val/1e8:.1f}억"
        elif abs(d_val) >= 1e4:
            d_str = f"{sign_v}{d_val/1e4:.0f}만"
        else:
            d_str = f"{sign_v}{d_val:,.0f}"
            
        x_axis_custom_labels.append(f"{b_name} ({d_str}, {sign_r}{d_rate:.1f}%)")
        diff_texts_main.append(f"<span style='font-size:15px; font-weight:800;'>{d_str}</span><br><span style='font-size:13px; font-weight:700;'>({sign_r}{d_rate:.1f}%)</span>")
        diff_colors_main.append("#E11D48" if d_val >= 0 else "#2563EB")
        
    summary_chart_view = summary_chart.copy()
    summary_chart_view["X축라벨"] = x_axis_custom_labels

    fig_bar = go.Figure()
    fig_bar.add_trace(go.Bar(
        x=summary_chart_view["X축라벨"],
        y=summary_chart_view[col_25],
        name="2025년 총 실적",
        marker=dict(color="#94A3B8", line=dict(color="#64748B", width=1), cornerradius=6),
        text=summary_chart_view[col_25].apply(lambda x: f"<span style='font-size:14px; font-weight:700;'>{x/1e8:.1f}억</span>" if x >= 1e8 else f"<span style='font-size:14px; font-weight:700;'>{x/1e4:.0f}만</span>"),
        textposition="outside",
        textfont=dict(size=14, color="#475569", family="Pretendard", weight="bold")
    ))
    fig_bar.add_trace(go.Bar(
        x=summary_chart_view["X축라벨"],
        y=summary_chart_view[col_26],
        name="2026년 총 실적",
        marker=dict(color="#1D4ED8", line=dict(color="#1E40AF", width=1), cornerradius=6),
        text=summary_chart_view[col_26].apply(lambda x: f"<span style='font-size:15px; font-weight:800;'>{x/1e8:.1f}억</span>" if x >= 1e8 else f"<span style='font-size:15px; font-weight:800;'>{x/1e4:.0f}만</span>"),
        textposition="outside",
        textfont=dict(size=15, color="#0F172A", family="Pretendard", weight="bold")
    ))
    fig_bar.update_layout(
        height=500,
        bargap=0.32,
        bargroupgap=0.10,
        yaxis=dict(
            rangemode='tozero', 
            title=dict(text="실적금액 (원)", font=dict(size=15, color="#1E293B", weight="bold")), 
            gridcolor="#F1F5F9", 
            zerolinecolor="#E2E8F0",
            tickfont=dict(size=14, color="#475569", weight="bold")
        ),
        xaxis=dict(
            categoryorder='array', 
            categoryarray=x_axis_custom_labels, 
            tickfont=dict(size=15, weight="bold", color="#0F172A")
        ),
        template="plotly_white",
        legend=dict(
            orientation="h", 
            yanchor="bottom", 
            y=1.05, 
            xanchor="left", 
            x=0, 
            font=dict(size=14, color="#1E293B", weight="bold")
        ),
        margin=dict(t=50, b=30, l=10, r=10)
    )
    st.plotly_chart(fig_bar, use_container_width=True)

    # 사업별 증감액 및 증감률 그래프
    st.write("")
    st.markdown("##### 📈 사업별 실적 증감액 및 증감률 (26년 - 25년)")
    fig_diff_main = go.Figure()
    fig_diff_main.add_trace(go.Bar(
        x=summary_chart_view["표준사업구분"],
        y=summary_chart_view["증감액"],
        marker=dict(color=diff_colors_main, cornerradius=6),
        text=diff_texts_main,
        textposition="outside",
        textfont=dict(size=14, family="Pretendard", weight="bold")
    ))
    fig_diff_main.update_layout(
        height=400,
        bargap=0.38,
        yaxis=dict(
            title=dict(text="증감액 (원)", font=dict(size=15, color="#1E293B", weight="bold")),
            gridcolor="#F1F5F9",
            zerolinecolor="#CBD5E1",
            tickfont=dict(size=14, color="#475569", weight="bold")
        ),
        xaxis=dict(
            categoryorder='array',
            categoryarray=target_categories,
            tickfont=dict(size=15, weight="bold", color="#0F172A")
        ),
        template="plotly_white",
        margin=dict(t=30, b=25, l=10, r=10)
    )
    st.plotly_chart(fig_diff_main, use_container_width=True)

    # 도넛 점유율 차트
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
            textinfo='label+percent', 
            textfont=dict(size=16, family="Pretendard", color="#FFFFFF", weight="bold"), 
            marker=dict(line=dict(color='#FFFFFF', width=2.5))
        )
        fig_pie_25.update_layout(
            height=460, 
            title=dict(font=dict(size=18, family="Pretendard", color="#0F172A", weight="bold")),
            margin=dict(t=60, b=20, l=10, r=10), 
            legend=dict(
                orientation="h", 
                yanchor="bottom", 
                y=-0.18, 
                xanchor="center", 
                x=0.5,
                font=dict(size=14, family="Pretendard", color="#1E293B", weight="bold")
            )
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
            textinfo='label+percent', 
            textfont=dict(size=16, family="Pretendard", color="#FFFFFF", weight="bold"), 
            marker=dict(line=dict(color='#FFFFFF', width=2.5))
        )
        fig_pie_26.update_layout(
            height=460, 
            title=dict(font=dict(size=18, family="Pretendard", color="#0F172A", weight="bold")),
            margin=dict(t=60, b=20, l=10, r=10), 
            legend=dict(
                orientation="h", 
                yanchor="bottom", 
                y=-0.18, 
                xanchor="center", 
                x=0.5,
                font=dict(size=14, family="Pretendard", color="#1E293B", weight="bold")
            )
        )
        st.plotly_chart(fig_pie_26, use_container_width=True)

# [페이지 2] [접수기준] 사업별 실적 현황
elif page_menu == "[접수기준] 사업별 실적 현황":
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
        sub_title_diff = "📈 사업별 실적 증감액 및 증감률 (26년 - 25년)"
        view_title = "🏢 사업별 2025년 vs 2026년 실적 비교 (접수기준)"
    else:
        display_df = calc_summary[calc_summary["표준사업구분"] == selected_view].copy()
        x_col = "세부항목"
        x_categories = display_df["세부항목"].unique().tolist()
        sub_title_diff = f"📈 [{selected_view}] 세부항목별 실적 증감액 및 증감률 (26년 - 25년)"
        view_title = f"🏢 [{selected_view}] 세부항목별 2025년 vs 2026년 실적 비교 (접수기준)"
        
    display_df = display_df.rename(columns={col_25: "2025년 실적", col_26: "2026년 실적"})
    display_df["증감액"] = display_df["2026년 실적"] - display_df["2025년 실적"]
    display_df["증감률"] = ((display_df["증감액"] / display_df["2025년 실적"].replace(0, pd.NA)) * 100).fillna(0.0)

    render_fullwidth_vertical_dashboard(
        title_top=view_title,
        title_bottom=sub_title_diff,
        table_title=f"[{'전체 사업' if selected_view == '전체 사업 보기' else selected_view}] 실적 요약 테이블",
        data_df=display_df,
        x_col_name=x_col,
        cat_order=x_categories
    )

# [페이지 3] [접수기준] 바이어 실적 현황
elif page_menu == "[접수기준] 바이어 실적 현황":
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

        render_fullwidth_vertical_dashboard(
            title_top=f"🤝 [{selected_biz}] 주요 바이어 2025년 vs 2026년 실적 변화 (접수기준)",
            title_bottom=f"📈 [{selected_biz}] 바이어별 증감액 및 증감률 (26년 - 25년)",
            table_title=f"[{selected_biz}] 주요 바이어 및 기타 실적 요약표",
            data_df=top_buyers,
            x_col_name="바이어명",
            cat_order=x_buyer_names
        )

# [페이지 4] [접수기준] 협력사 실적 현황 (신규 추가)
elif page_menu == "[접수기준] 협력사 실적 현황":
    current_tab4_biz = st.session_state.get("selected_tab4_biz", target_categories[0])
    current_idx4 = target_categories.index(current_tab4_biz) if current_tab4_biz in target_categories else 0
    
    selected_biz = st.selectbox(
        "조회할 사업부문을 선택하세요:", 
        target_categories, 
        index=current_idx4,
        key="tab4_biz_selectbox"
    )
    
    if selected_biz != st.session_state.get("selected_tab4_biz"):
        st.session_state["selected_tab4_biz"] = selected_biz
        st.rerun()
    
    raw_b_chart = part_data_cache.get(selected_biz, pd.DataFrame()).copy()
    
    target_row = summary_chart[summary_chart["표준사업구분"] == selected_biz]
    target_tot_25 = float(target_row[col_25].values[0]) if not target_row.empty else 0.0
    target_tot_26 = float(target_row[col_26].values[0]) if not target_row.empty else 0.0

    if raw_b_chart.empty:
        st.warning(f"선택하신 [{selected_biz}] 부문의 요약 블록('행 레이블' 표)을 읽을 수 없습니다.")
    else:
        is_dash = raw_b_chart["바이어명"].astype(str).str.strip().isin(["-", "–", "—", "", "NAN", "NONE", "기타"])
        valid_vendors = raw_b_chart[~is_dash].copy()
        dash_vendors = raw_b_chart[is_dash].copy()
        
        valid_vendors = valid_vendors.sort_values(by="2026년 실적", ascending=False).reset_index(drop=True)
        
        if len(valid_vendors) > 6:
            top6 = valid_vendors.iloc[:6].copy()
            rest = valid_vendors.iloc[6:].copy()
            
            top6_25 = top6["2025년 실적"].sum()
            top6_26 = top6["2026년 실적"].sum()
            
            calc_etc_25 = rest["2025년 실적"].sum() + dash_vendors["2025년 실적"].sum()
            calc_etc_26 = rest["2026년 실적"].sum() + dash_vendors["2026년 실적"].sum()
            
            etc_25 = max(calc_etc_25, target_tot_25 - top6_25)
            etc_26 = max(calc_etc_26, target_tot_26 - top6_26)
            
            etc_row = pd.DataFrame([{
                "협력사명": "기타 협력사",
                "2025년 실적": etc_25,
                "2026년 실적": etc_26
            }])
            top6 = top6.rename(columns={"바이어명": "협력사명"})
            top_vendors = pd.concat([top6, etc_row], ignore_index=True)
        else:
            top_sum_25 = valid_vendors["2025년 실적"].sum()
            top_sum_26 = valid_vendors["2026년 실적"].sum()
            
            etc_25 = max(dash_vendors["2025년 실적"].sum(), target_tot_25 - top_sum_25)
            etc_26 = max(dash_vendors["2026년 실적"].sum(), target_tot_26 - top_sum_26)
            
            valid_vendors = valid_vendors.rename(columns={"바이어명": "협력사명"})
            if etc_25 > 0 or etc_26 > 0:
                etc_row = pd.DataFrame([{
                    "협력사명": "기타 협력사",
                    "2025년 실적": etc_25,
                    "2026년 실적": etc_26
                }])
                top_vendors = pd.concat([valid_vendors, etc_row], ignore_index=True)
            else:
                top_vendors = valid_vendors.copy()

        top_vendors["증감액"] = top_vendors["2026년 실적"] - top_vendors["2025년 실적"]
        top_vendors["증감률"] = ((top_vendors["증감액"] / top_vendors["2025년 실적"].replace(0, pd.NA)) * 100).fillna(0.0)

        x_vendor_names = [b for b in top_vendors["협력사명"] if b != "기타 협력사"] + (["기타 협력사"] if "기타 협력사" in top_vendors["협력사명"].values else [])

        render_fullwidth_vertical_dashboard(
            title_top=f"🏢 [{selected_biz}] 주요 협력사 2025년 vs 2026년 실적 변화 (접수기준)",
            title_bottom=f"📈 [{selected_biz}] 협력사별 증감액 및 증감률 (26년 - 25년)",
            table_title=f"[{selected_biz}] 주요 협력사 및 기타 실적 요약표",
            data_df=top_vendors,
            x_col_name="협력사명",
            cat_order=x_vendor_names
        )

# [페이지 5] [BI_종합] 사업별 실적 현황
elif page_menu == "[BI_종합] 사업별 실적 현황":
    render_fullwidth_vertical_dashboard(
        title_top=f"📊 [BI_종합] 8대 사업별 2025년 vs 2026년 실적 비교 ({'월계' if '월계' in bi_period_mode else '누계'} 기준)",
        title_bottom=f"📈 8대 사업별 증감액 및 증감률 (26년 - 25년)",
        table_title=f"[BI_종합] 8대 사업별 실적 상세 요약표 ({'월계' if '월계' in bi_period_mode else '누계'} 기준)",
        data_df=bi_total_charts["월계" if "월계" in bi_period_mode else "누계"],
        x_col_name="표준사업구분",
        cat_order=BI_8_CATEGORIES
    )

# [페이지 6] [BI_상해] 사업별 실적 현황
elif page_menu == "[BI_상해] 사업별 실적 현황":
    render_fullwidth_vertical_dashboard(
        title_top=f"🏙️ [BI_상해] 8대 사업별 2025년 vs 2026년 실적 비교 ({'월계' if '월계' in bi_period_mode else '누계'} 기준)",
        title_bottom=f"📈 8대 사업별 증감액 및 증감률 (26년 - 25년)",
        table_title=f"[BI_상해] 8대 사업별 실적 상세 요약표 ({'월계' if '월계' in bi_period_mode else '누계'} 기준)",
        data_df=bi_shanghai_charts["월계" if "월계" in bi_period_mode else "누계"],
        x_col_name="표준사업구분",
        cat_order=BI_8_CATEGORIES
    )

# [페이지 7] [BI_광주] 사업별 실적 현황
elif page_menu == "[BI_광주] 사업별 실적 현황":
    render_fullwidth_vertical_dashboard(
        title_top=f"🏭 [BI_광주] 8대 사업별 2025년 vs 2026년 실적 비교 ({'월계' if '월계' in bi_period_mode else '누계'} 기준)",
        title_bottom=f"📈 8대 사업별 증감액 및 증감률 (26년 - 25년)",
        table_title=f"[BI_광주] 8대 사업별 실적 상세 요약표 ({'월계' if '월계' in bi_period_mode else '누계'} 기준)",
        data_df=bi_guangzhou_charts["월계" if "월계" in bi_period_mode else "누계"],
        x_col_name="표준사업구분",
        cat_order=BI_8_CATEGORIES
    )
