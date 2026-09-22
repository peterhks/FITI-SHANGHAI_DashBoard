import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import io

# =========================================================
# 1. 화면 기본 설정 및 디자인 스타일
# =========================================================
st.set_page_config(
    page_title="FITI SHANGHAI Performance Analysis",
    layout="wide",
    initial_sidebar_state="expanded"
)

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

# =========================================================
# 2. 다국어 텍스트 사전 (한국어, 중국어, 영어)
# =========================================================
LANG_DICT = {
    "한국어": {
        "sys_title": "상해지사 실적 종합 분석 시스템",
        "sys_sub": "상해지사 사업 실적 및 분석 시스템 | 상해지사 사업팀",
        "data_mgmt": "📁 데이터 관리",
        "admin_upload": "공용 엑셀 파일 업로드 (관리자 전용)",
        "admin_caption": "💡 엑셀 파일을 교체하려면 하단 'BI 관리자 모드'로 로그인하세요.",
        "sync_success": "✅ 서버 공용 파일 및 세션 동기화 완료!",
        "shared_file_info": "📂 서버 공용 최신 파일 연동 중",
        "file_not_found": "분석할 엑셀 파일을 찾을 수 없습니다. 관리자 모드로 로그인하여 파일을 업로드해 주세요.",
        "page_select": "📑 분석 페이지 선택",
        "cat_select": "📌 카테고리 선택",
        "period_select": "⏱️ [BI] 실적 기간 선택",
        "auth_title": "🔒 BI 실적 보안 인증",
        "auth_input": "열람 비밀번호 입력:",
        "auth_fail": "비밀번호가 일치하지 않습니다.",
        "auth_success": "🔓 BI 관리자 모드 활성화됨",
        "logout_btn": "BI 잠금 (로그아웃)",
        "kpi_25": "📅 25년 총 실적",
        "kpi_26": "🚀 26년 총 실적",
        "kpi_diff": "📈 실적 증감액",
        "kpi_rate": "📊 증감 퍼센트",
        "kpi_diff_sub": "전년 대비 실적차",
        "kpi_rate_sub": "전년 대비 성장률",
        "pie_title_25": "2025년 사업별 실적 비중",
        "pie_title_26": "2026년 사업별 실적 비중",
        "unit": "원",
        "font_family": "'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif",
        "pages": {
            "[접수기준] 종합 실적 현황": "[접수기준] 종합 실적 현황",
            "[접수기준] 사업별 실적 현황": "[접수기준] 사업별 실적 현황",
            "[접수기준] 바이어 실적 현황": "[접수기준] 바이어 실적 현황",
            "[접수기준] 협력사 실적 현황": "[접수기준] 협력사 실적 현황",
            "[BI_종합] 사업별 실적 현황": "[BI_종합] 사업별 실적 현황",
            "[BI_상해] 사업별 실적 현황": "[BI_상해] 사업별 실적 현황",
            "[BI_광주] 사업별 실적 현황": "[BI_광주] 사업별 실적 현황",
        },
        "periods": {
            "전체 총계 누계": "전체 총계 누계",
            "사업 소계 누계": "사업 소계 누계",
            "사업 소계 월계": "사업 소계 월계"
        },
        "categories_map": {
            "글로벌 바이어": "글로벌 바이어",
            "패션잡화": "패션잡화",
            "GB": "GB",
            "제품평가": "제품평가"
        }
    },
    "中文 (중국어)": {
        "sys_title": "上海分公司业绩综合分析系统",
        "sys_sub": "上海分公司业务业绩及分析系统 | 上海分公司业务团队",
        "data_mgmt": "📁 数据管理",
        "admin_upload": "上传公共Excel文件 (仅限管理员)",
        "admin_caption": "💡 如需更换Excel文件，请登录底部的“BI管理员模式”。",
        "sync_success": "✅ 服务器公共文件及会话同步完成！",
        "shared_file_info": "📂 服务器公共最新文件同步中",
        "file_not_found": "未找到要分析的Excel文件。请登录管理员模式上传。",
        "page_select": "📑 选择分析页面",
        "cat_select": "📌 选择类别",
        "period_select": "⏱️ [BI] 业绩期间选择",
        "auth_title": "🔒 BI 业绩安全验证",
        "auth_input": "请输入查看密码:",
        "auth_fail": "密码不正确。",
        "auth_success": "🔓 BI 管理员模式已激活",
        "logout_btn": "锁定 BI (登出)",
        "kpi_25": "📅 25年总业绩",
        "kpi_26": "🚀 26年总业绩",
        "kpi_diff": "📈 业绩增减额",
        "kpi_rate": "📊 增减百分比",
        "kpi_diff_sub": "较去年业绩差额",
        "kpi_rate_sub": "较去年增长率",
        "pie_title_25": "2025年各业务业绩占比",
        "pie_title_26": "2026年各业务业绩占比",
        "unit": "韩元",
        "font_family": "'SimHei', '黑体', sans-serif",
        "pages": {
            "[접수기준] 종합 실적 현황": "[接收基准] 综合业绩现状",
            "[접수기준] 사업별 실적 현황": "[接收基准] 各业务业绩现状",
            "[접수기준] 바이어 실적 현황": "[接收基准] 各买家业绩现状",
            "[접수기준] 협력사 실적 현황": "[接收基准] 各合作社业绩现状",
            "[BI_종합] 사업별 실적 현황": "[BI_综合] 各业务业绩现状",
            "[BI_상해] 사업별 실적 현황": "[BI_上海] 各业务业绩现状",
            "[BI_광주] 사업별 실적 현황": "[BI_光州] 各业务业绩现状",
        },
        "periods": {
            "전체 총계 누계": "全体总计累计",
            "사업 소계 누계": "业务小计累计",
            "사업 소계 월계": "业务小计月度"
        },
        "categories_map": {
            "글로벌 바이어": "全球买家",
            "패션잡화": "时尚杂货",
            "GB": "GB标准",
            "제품평가": "产品评价"
        }
    },
    "English (영어)": {
        "sys_title": "Shanghai Branch Performance Analysis System",
        "sys_sub": "Shanghai Branch Business Performance & Analysis System | Business Team",
        "data_mgmt": "📁 Data Management",
        "admin_upload": "Upload Public Excel (Admin Only)",
        "admin_caption": "💡 To replace Excel, login to 'BI Admin Mode' below.",
        "sync_success": "✅ Server public file & session synced!",
        "shared_file_info": "📂 Server Public Latest File Linked",
        "file_not_found": "Excel file not found. Please login as admin to upload.",
        "page_select": "📑 Select Page",
        "cat_select": "📌 Select Category",
        "period_select": "⏱️ [BI] Period Select",
        "auth_title": "🔒 BI Security Auth",
        "auth_input": "Enter password:",
        "auth_fail": "Incorrect password.",
        "auth_success": "🔓 BI Admin Mode Active",
        "logout_btn": "Lock BI (Logout)",
        "kpi_25": "📅 '25 Total Performance",
        "kpi_26": "🚀 '26 Total Performance",
        "kpi_diff": "📈 Performance Diff",
        "kpi_rate": "📊 Growth Rate",
        "kpi_diff_sub": "YoY Performance Gap",
        "kpi_rate_sub": "YoY Growth Rate",
        "pie_title_25": "2025 Performance Share by Business",
        "pie_title_26": "2026 Performance Share by Business",
        "unit": "KRW",
        "font_family": "'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif",
        "pages": {
            "[접수기준] 종합 실적 현황": "[Receipt Basis] Overall Performance",
            "[접수기준] 사업별 실적 현황": "[Receipt Basis] Performance by Business",
            "[접수기준] 바이어 실적 현황": "[Receipt Basis] Performance by Buyer",
            "[접수기준] 협력사 실적 현황": "[Receipt Basis] Performance by Vendor",
            "[BI_종합] 사업별 실적 현황": "[BI_Total] Performance by Business",
            "[BI_상해] 사업별 실적 현황": "[BI_Shanghai] Performance by Business",
            "[BI_광주] 사업별 실적 현황": "[BI_Gwangju] Performance by Business",
        },
        "periods": {
            "전체 총계 누계": "Total Cumulative",
            "사업 소계 누계": "Business Subtotal Cumulative",
            "사업 소계 월계": "Business Subtotal Monthly"
        },
        "categories_map": {
            "글로벌 바이어": "Global Buyer",
            "패션잡화": "Fashion & Misc",
            "GB": "China GB",
            "제품평가": "Product Inspection"
        }
    }
}

# =========================================================
# 3. 사이드바 언어 선택 및 쿼리 연동
# =========================================================
query_params = st.query_params

if "lang" in query_params and query_params["lang"] in LANG_DICT:
    st.session_state["selected_lang"] = query_params["lang"]

if "selected_lang" not in st.session_state:
    st.session_state["selected_lang"] = "한국어"

def on_lang_change():
    st.query_params["lang"] = st.session_state["lang_selectbox"]

st.sidebar.markdown(f"### 🌐 언어 설정 / 语言设置 / Language")
selected_lang = st.sidebar.selectbox(
    "표시 언어 선택", 
    ["한국어", "中文 (중국어)", "English (영어)"],
    index=["한국어", "中文 (중국어)", "English (영어)"].index(st.session_state["selected_lang"]),
    key="lang_selectbox",
    on_change=on_lang_change,
    label_visibility="collapsed"
)

st.session_state["selected_lang"] = selected_lang
t = LANG_DICT[selected_lang]
current_font = t["font_family"]

st.markdown(f"""
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    html, body, [class*="css"] {{
        font-family: {current_font} !important;
    }}
    
    .fiti-header {{
        background: linear-gradient(135deg, #002B5C 0%, #003876 100%);
        padding: 22px 28px;
        border-radius: 10px;
        display: flex;
        align-items: center;
        gap: 22px;
        color: #FFFFFF;
        margin-bottom: 22px;
        box-shadow: 0 4px 14px rgba(0, 43, 92, 0.18);
    }}
    .fiti-logo-text {{
        font-size: 28px;
        font-weight: 900;
        letter-spacing: -0.5px;
        border-right: 1.5px solid rgba(255, 255, 255, 0.25);
        padding-right: 22px;
    }}
    .fiti-title-main {{
        font-size: 21px;
        font-weight: 800;
        margin-bottom: 4px;
        letter-spacing: -0.3px;
    }}
    .fiti-title-sub {{
        font-size: 13px;
        color: #D0E1FD;
        font-weight: 400;
    }}

    .kpi-card {{
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 20px 22px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04), 0 2px 4px -2px rgba(0, 0, 0, 0.04);
        transition: transform 0.15s ease, box-shadow 0.15s ease;
        border-top: 4px solid #CBD5E1;
    }}
    .kpi-card:hover {{
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08);
    }}
    .kpi-title {{
        font-size: 13px;
        font-weight: 600;
        color: #64748B;
        margin-bottom: 8px;
    }}
    .kpi-num {{
        font-size: 26px;
        font-weight: 800;
        color: #0F172A;
        letter-spacing: -0.5px;
    }}
    .kpi-sub {{
        font-size: 12px;
        color: #94A3B8;
        margin-top: 6px;
    }}
    .kpi-badge {{
        display: inline-block;
        padding: 3px 8px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 700;
        margin-top: 6px;
    }}

    .sidebar-card-btn {{
        display: block;
        width: 100%;
        border-radius: 8px;
        text-align: center;
        font-weight: 700;
        font-size: 14px;
        padding: 11px 14px;
        margin-bottom: 5px;
        border: 1.5px solid #CBD5E1;
        background-color: #F8FAFC;
        color: #0F172A;
        text-decoration: none;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04);
        transition: all 0.15s ease;
    }}
    .sidebar-card-btn:hover {{
        border-color: #003876;
        background-color: #E2E8F0;
        color: #002B5C;
    }}
    .sidebar-card-btn-active {{
        display: block;
        width: 100%;
        border-radius: 8px;
        text-align: center;
        font-weight: 800;
        font-size: 14px;
        padding: 11px 14px;
        margin-bottom: 5px;
        border: 1.5px solid #001E3D;
        background-color: #003876;
        color: #FFFFFF !important;
        text-decoration: none;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }}
</style>
""", unsafe_allow_html=True)

# =========================================================
# 4. 상단 공식 배너
# =========================================================
st.markdown(f"""
<div class="fiti-header">
    <div class="fiti-logo-text">FITI</div>
    <div>
        <div class="fiti-title-main">{t["sys_title"]}</div>
        <div class="fiti-title-sub">{t["sys_sub"]}</div>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================
# 5. 보안 인증 및 관리자 모드 전용 업로드 엔진
# =========================================================
EXCEL_FILE = "performance_최신.xlsx"
os.makedirs("downloads", exist_ok=True)
LOCAL_EXCEL_PATH = os.path.join("downloads", EXCEL_FILE)

st.sidebar.markdown(f"### {t['data_mgmt']}")

if "bi_authorized" not in st.session_state:
    st.session_state["bi_authorized"] = False

def check_bi_password():
    pw_val = st.session_state.get("bi_pw_input", "")
    if pw_val == "fiti1965":
        st.session_state["bi_authorized"] = True
        st.query_params["auth"] = "true"
    else:
        st.session_state["bi_authorized"] = False
        st.sidebar.error(t["auth_fail"])

if st.session_state["bi_authorized"]:
    uploaded_file = st.sidebar.file_uploader(t["admin_upload"], type=["xlsx", "csv"])
    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        with open(LOCAL_EXCEL_PATH, "wb") as f:
            f.write(file_bytes)
        st.session_state["persistent_file_bytes"] = file_bytes
        st.session_state["persistent_file_name"] = uploaded_file.name
        st.cache_data.clear()
        st.sidebar.success(t["sync_success"])
        st.rerun()
else:
    st.sidebar.caption(t["admin_caption"])

if "persistent_file_bytes" in st.session_state:
    raw_bytes = st.session_state["persistent_file_bytes"]
    st.sidebar.success(t["shared_file_info"])
elif os.path.exists(LOCAL_EXCEL_PATH):
    with open(LOCAL_EXCEL_PATH, "rb") as f:
        raw_bytes = f.read()
    st.session_state["persistent_file_bytes"] = raw_bytes
    st.sidebar.info(t["shared_file_info"])
elif os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        raw_bytes = f.read()
    st.session_state["persistent_file_bytes"] = raw_bytes
    st.sidebar.info(t["shared_file_info"])
else:
    raw_bytes = None

if not raw_bytes:
    st.warning(t["file_not_found"])
    st.stop()

def clean_series(series):
    cleaned = series.astype(str).str.replace(',', '').str.replace('₩', '').str.strip()
    cleaned = cleaned.replace(['-', '–', '—', 'nan', 'NaN', 'None', ''], '0')
    return pd.to_numeric(cleaned, errors='coerce').fillna(0)

@st.cache_data
def get_excel_sheets(file_bytes_val):
    stream = io.BytesIO(file_bytes_val)
    excel_obj = pd.ExcelFile(stream)
    all_sheets = excel_obj.sheet_names
    sheet_dict = {s.strip().lower().replace(" ", "").replace("_", ""): s for s in all_sheets}
    return all_sheets, sheet_dict

sheet_names, sheet_dict = get_excel_sheets(raw_bytes)

def get_sheet_by_keyword(keywords):
    for s_clean, orig_name in sheet_dict.items():
        if all(k.lower().replace(" ", "").replace("_", "") in s_clean for k in keywords):
            return orig_name
    return None

# =========================================================
# 6. '종합' 시트 파서
# =========================================================
@st.cache_data
def parse_summary_data(file_bytes_val):
    stream = io.BytesIO(file_bytes_val)
    summary_sheet_name = get_sheet_by_keyword(["종합"]) or sheet_names[0]
    raw_summary = pd.read_excel(stream, sheet_name=summary_sheet_name, header=None)

    h_idx = 0
    for idx, row in raw_summary.iterrows():
        r_text = "".join(row.dropna().astype(str).tolist())
        if "구분" in r_text and any(k in r_text for k in ["합계", "25", "26"]):
            h_idx = idx
            break

    stream.seek(0)
    df_summary = pd.read_excel(stream, sheet_name=summary_sheet_name, skiprows=h_idx)

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

    return summary_chart, calc_summary, col_25, col_26, target_categories

summary_chart, calc_summary, col_25, col_26, target_categories = parse_summary_data(raw_bytes)

# =========================================================
# 7. 세부 파트 및 지사별 파서
# =========================================================
PART_SHEET_MAPPINGS = {
    "패션잡화": [["kc"]],
    "GB": [["gb"]],
    "글로벌 바이어": [["global", "1"], ["global", "2"]],
    "제품평가": [["inspection", "원단"], ["inspection", "가먼트"]]
}

@st.cache_data
def extract_pivot_block(file_bytes_val, sheet_name):
    stream = io.BytesIO(file_bytes_val)
    raw = pd.read_excel(stream, sheet_name=sheet_name, header=None)
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

@st.cache_data
def extract_vendor_data_from_sheet(file_bytes_val, sheet_name):
    stream = io.BytesIO(file_bytes_val)
    raw = pd.read_excel(stream, sheet_name=sheet_name, header=None)
    h_idx, buyer_col_idx, vendor_col_idx = None, None, None
    
    for r_i in range(min(15, len(raw))):
        row_vals = [str(x).strip().replace(" ", "") for x in raw.iloc[r_i].tolist()]
        for c_i, val in enumerate(row_vals):
            if "업체명" in val or "협력사" in val:
                vendor_col_idx = c_i
            elif "바이어" in val:
                buyer_col_idx = c_i
        if vendor_col_idx is not None:
            h_idx = r_i
            break
            
    if h_idx is None or vendor_col_idx is None:
        return pd.DataFrame(columns=["협력사명", "바이어명", "2025년 실적", "2026년 실적"])
        
    stream.seek(0)
    df_raw = pd.read_excel(stream, sheet_name=sheet_name, skiprows=h_idx)
    v_col = df_raw.columns[vendor_col_idx]
    b_col = df_raw.columns[buyer_col_idx] if buyer_col_idx is not None and buyer_col_idx < len(df_raw.columns) else None
    
    c25, c26 = None, None
    for c in df_raw.columns:
        c_str = str(c).replace(" ", "")
        if "25" in c_str and ("합계" in c_str or "실적" in c_str or "총" in c_str):
            c25 = c
        elif "26" in c_str and ("합계" in c_str or "실적" in c_str or "총" in c_str):
            c26 = c
            
    if not c25 or not c26:
        for c in df_raw.columns:
            c_str = str(c).replace(" ", "")
            if "25" in c_str and c25 is None:
                c25 = c
            elif "26" in c_str and c26 is None:
                c26 = c
                
    if not c25 or not c26:
        return pd.DataFrame(columns=["협력사명", "바이어명", "2025년 실적", "2026년 실적"])
        
    df_clean = df_raw.dropna(subset=[v_col]).copy()
    df_clean = df_clean[~df_clean[v_col].astype(str).str.contains(r"소계|합계|TOTAL|총계", regex=True, na=False)].copy()
    
    res = pd.DataFrame()
    res["협력사명"] = df_clean[v_col].astype(str).str.strip()
    res["바이어명"] = df_clean[b_col].astype(str).str.strip() if b_col else "기본"
    res["2025년 실적"] = clean_series(df_clean[c25])
    res["2026년 실적"] = clean_series(df_clean[c26])
    
    return res[res["협력사명"] != ""]

part_data_cache = {}
vendor_data_cache = {}

for cat in target_categories:
    keywords_list = PART_SHEET_MAPPINGS.get(cat, [])
    b_dfs, v_dfs = [], []
    for k_words in keywords_list:
        sheet_n = get_sheet_by_keyword(k_words)
        if sheet_n:
            b_df = extract_pivot_block(raw_bytes, sheet_n)
            if not b_df.empty:
                b_dfs.append(b_df)
            v_df = extract_vendor_data_from_sheet(raw_bytes, sheet_n)
            if not v_df.empty:
                v_dfs.append(v_df)
    
    if b_dfs:
        comb_b = pd.concat(b_dfs, ignore_index=True)
        part_data_cache[cat] = comb_b.groupby("바이어명", as_index=False)[["2025년 실적", "2026년 실적"]].sum()
    else:
        part_data_cache[cat] = pd.DataFrame(columns=["바이어명", "2025년 실적", "2026년 실적"])
        
    if v_dfs:
        vendor_data_cache[cat] = pd.concat(v_dfs, ignore_index=True)
    else:
        vendor_data_cache[cat] = pd.DataFrame(columns=["협력사명", "바이어명", "2025년 실적", "2026년 실적"])

@st.cache_data
def parse_bi_sheet_by_type(file_bytes_val, branch_name="종합"):
    stream = io.BytesIO(file_bytes_val)
    excel_obj = pd.ExcelFile(stream)
    all_sheets = excel_obj.sheet_names
    
    target_s_name = None
    search_key = branch_name.strip().lower().replace(" ", "").replace("_", "")
    
    for s_orig in all_sheets:
        s_clean = s_orig.strip().lower().replace(" ", "").replace("_", "")
        if search_key == "광주" and s_clean == "bi광주":
            target_s_name = s_orig
            break
        elif search_key == "상해" and s_clean == "bi상해":
            target_s_name = s_orig
            break
        elif search_key == "종합" and s_clean in ["bi종합", "bi"]:
            target_s_name = s_orig
            break
            
    if not target_s_name:
        for s_orig in all_sheets:
            s_clean = s_orig.strip().lower().replace("_", "").replace(" ", "")
            if search_key == "광주" and "광주" in s_clean and "상해" not in s_clean:
                target_s_name = s_orig
                break
            elif search_key == "상해" and "상해" in s_clean and "광주" not in s_clean:
                target_s_name = s_orig
                break
            elif search_key == "종합" and "bi" in s_clean and "광주" not in s_clean and "상해" not in s_clean:
                target_s_name = s_orig
                break

    empty_kpi = {"25": 0, "26": 0, "diff": 0, "rate": 0.0}
    def_kpi = {
        "전체 총계 누계": empty_kpi,
        "사업 소계 누계": empty_kpi,
        "사업 소계 월계": empty_kpi
    }
    
    empty_df = pd.DataFrame({
        "표준사업구분": BI_8_CATEGORIES,
        "2025년 실적": [0]*8,
        "2026년 실적": [0]*8,
        "증감률": [0.0]*8
    })
    def_chart = {"누계": empty_df, "월계": empty_df}

    if not target_s_name:
        return def_kpi, def_chart

    stream.seek(0)
    raw = pd.read_excel(stream, sheet_name=target_s_name, header=None)

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
        "전체 총계 누계": extract_row_vals(total_r_idx, c_c25, c_c26, c_rate, empty_kpi),
        "사업 소계 누계": extract_row_vals(subtotal_r_idx, c_c25, c_c26, c_rate, empty_kpi),
        "사업 소계 월계": extract_row_vals(subtotal_r_idx, m_c25, m_c26, m_rate, m_rate if m_rate else c_rate)
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

bi_total_kpi, bi_total_charts = parse_bi_sheet_by_type(raw_bytes, "종합")
bi_shanghai_kpi, bi_shanghai_charts = parse_bi_sheet_by_type(raw_bytes, "상해")
bi_guangzhou_kpi, bi_guangzhou_charts = parse_bi_sheet_by_type(raw_bytes, "광주")

# =========================================================
# 9. 사이드바 네비게이션 및 다국어 상태 유지 링크 연동
# =========================================================
st.sidebar.markdown(f"### {t['page_select']}")

base_pages_keys = [
    "[접수기준] 종합 실적 현황",
    "[접수기준] 사업별 실적 현황",
    "[접수기준] 바이어 실적 현황",
    "[접수기준] 협력사 실적 현황"
]

if "auth" in query_params and query_params["auth"] == "true":
    st.session_state["bi_authorized"] = True

if "bi_authorized" not in st.session_state:
    st.session_state["bi_authorized"] = False

def check_bi_password():
    pw_val = st.session_state.get("bi_pw_input", "")
    if pw_val == "fiti1965":
        st.session_state["bi_authorized"] = True
        st.query_params["auth"] = "true"
    else:
        st.session_state["bi_authorized"] = False
        st.sidebar.error(t["auth_fail"])

if st.session_state["bi_authorized"]:
    bi_pages_keys = [
        "[BI_종합] 사업별 실적 현황",
        "[BI_상해] 사업별 실적 현황",
        "[BI_광주] 사업별 실적 현황"
    ]
else:
    bi_pages_keys = []

all_pages_keys = base_pages_keys + bi_pages_keys

if "current_page" not in st.session_state:
    st.session_state["current_page"] = all_pages_keys[0]

if "current_page" not in all_pages_keys:
    st.session_state["current_page"] = all_pages_keys[0]

if "page" in query_params:
    p_param = query_params["page"]
    if p_param in all_pages_keys:
        st.session_state["current_page"] = p_param

st.sidebar.markdown(f"##### {t['cat_select']}")

auth_param_str = "&auth=true" if st.session_state["bi_authorized"] else ""
lang_param_str = f"&lang={selected_lang}"

for p_key in all_pages_keys:
    is_active = (st.session_state["current_page"] == p_key)
    btn_class = "sidebar-card-btn-active" if is_active else "sidebar-card-btn"
    display_name = t["pages"].get(p_key, p_key)
    
    card_link_html = f"""
    <a href="?page={p_key}{auth_param_str}{lang_param_str}" class="{btn_class}" target="_self">
        {display_name}
    </a>
    """
    st.sidebar.markdown(card_link_html, unsafe_allow_html=True)

page_menu = st.session_state["current_page"]

st.sidebar.markdown("---")

if not st.session_state["bi_authorized"]:
    st.sidebar.markdown(f"##### {t['auth_title']}")
    st.sidebar.text_input(
        t["auth_input"], 
        type="password", 
        key="bi_pw_input", 
        on_change=check_bi_password
    )
else:
    st.sidebar.markdown(f"##### {t['auth_success']}")
    if st.sidebar.button(t["logout_btn"], key="logout_btn_unique_99"):
        st.session_state["bi_authorized"] = False
        st.query_params.clear()
        st.rerun()

# =========================================================
# 10. 상단 종합 KPI 카드 및 다국어 렌더링
# =========================================================
card_unit = t["unit"]

# 💡 [안정화] display_period_name 변수를 상단에서 미리 정의하여 NameError 완전 차단
display_period_name = "전체 총계 누계"

if page_menu.startswith("[BI_"):
    st.sidebar.markdown("---")
    st.sidebar.markdown(f"### {t['period_select']}")
    
    bi_periods_keys = ["전체 총계 누계", "사업 소계 누계", "사업 소계 월계"]
    if "bi_period_mode" not in st.session_state:
        st.session_state["bi_period_mode"] = bi_periods_keys[0]
        
    period_query = query_params.get("period", None)
    if period_query in bi_periods_keys:
        st.session_state["bi_period_mode"] = period_query

    for bp_key in bi_periods_keys:
        is_p_active = (st.session_state["bi_period_mode"] == bp_key)
        p_class = "sidebar-card-btn-active" if is_p_active else "sidebar-card-btn"
        
        period_link_html = f"""
        <a href="?page={page_menu}&period={bp_key}{auth_param_str}{lang_param_str}" class="{p_class}" target="_self">
            {bp_key}
        </a>
        """
        st.sidebar.markdown(period_link_html, unsafe_allow_html=True)
            
    bi_period_mode = st.session_state["bi_period_mode"]
    display_period_name = t["periods"].get(bi_period_mode, bi_period_mode)
    
    if "광주" in page_menu:
        target_kpi_pack = bi_guangzhou_kpi
        active_bi_charts = bi_guangzhou_charts
        sub_prefix = "BI_광주"
    elif "상해" in page_menu:
        target_kpi_pack = bi_shanghai_kpi
        active_bi_charts = bi_shanghai_charts
        sub_prefix = "BI_상해"
    else:
        target_kpi_pack = bi_total_kpi
        active_bi_charts = bi_total_charts
        sub_prefix = "BI_종합"

    bi_pack = target_kpi_pack.get(bi_period_mode, target_kpi_pack["전체 총계 누계"])
    total_25 = float(bi_pack["25"])
    total_26 = float(bi_pack["26"])
    diff_val = float(bi_pack["diff"])
    diff_rate = float(bi_pack["rate"])
    card_sub_desc = f"{sub_prefix} [{display_period_name}]"
else:
    selected_view_for_card = "전체 사업 보기"
    if page_menu == "[접수기준] 사업별 실적 현황":
        if "selected_biz_view" not in st.session_state:
            st.session_state["selected_biz_view"] = "전체 사업 보기"
        selected_view_for_card = st.session_state["selected_biz_view"]
    elif page_menu in ["[접수기준] 바이어 실적 현황", "[접수기준] 협력사 실적 현황"]:
        card_key = "selected_tab3_biz" if page_menu == "[접수기준] 바이어 실적 현황" else "tab4_biz_select"
        if card_key not in st.session_state:
            st.session_state[card_key] = target_categories[0]
        selected_view_for_card = st.session_state[card_key]

    if selected_view_for_card != "전체 사업 보기" and selected_view_for_card in target_categories:
        target_row = summary_chart[summary_chart["표준사업구분"] == selected_view_for_card]
        total_25 = float(target_row[col_25].sum()) if not target_row.empty else 0.0
        total_26 = float(target_row[col_26].sum()) if not target_row.empty else 0.0
        display_cat_name = t["categories_map"].get(selected_view_for_card, selected_view_for_card)
        card_sub_desc = f"[{display_cat_name}] Total"
    else:
        total_25 = float(summary_chart[col_25].sum())
        total_26 = float(summary_chart[col_26].sum())
        card_sub_desc = "TOTAL Summary"

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
        <div class="kpi-title">{t["kpi_25"]}</div>
        <div class="kpi-num">{total_25:,.0f} <span style="font-size:14px; font-weight:normal; color:#64748B;">{card_unit}</span></div>
        <div class="kpi-sub">{card_sub_desc}</div>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown(f"""
    <div class="kpi-card" style="border-top-color: #003876;">
        <div class="kpi-title">{t["kpi_26"]}</div>
        <div class="kpi-num" style="color: #003876;">{total_26:,.0f} <span style="font-size:14px; font-weight:normal; color:#64748B;">{card_unit}</span></div>
        <div class="kpi-sub">{card_sub_desc}</div>
    </div>
    """, unsafe_allow_html=True)

with c3:
    st.markdown(f"""
    <div class="kpi-card" style="border-top-color: {diff_color};">
        <div class="kpi-title">{t["kpi_diff"]}</div>
        <div class="kpi-num" style="color: {diff_color};">{diff_sign}{diff_val:,.0f} <span style="font-size:14px; font-weight:normal; color:#64748B;">{card_unit}</span></div>
        <span class="kpi-badge" style="background-color: {badge_bg}; color: {diff_color};">{t["kpi_diff_sub"]}</span>
    </div>
    """, unsafe_allow_html=True)

with c4:
    st.markdown(f"""
    <div class="kpi-card" style="border-top-color: {diff_color};">
        <div class="kpi-title">{t["kpi_rate"]}</div>
        <div class="kpi-num" style="color: {diff_color};">{diff_sign}{diff_rate:0.1f}%</div>
        <span class="kpi-badge" style="background-color: {badge_bg}; color: {diff_color};">{t["kpi_rate_sub"]}</span>
    </div>
    """, unsafe_allow_html=True)

st.write("")
st.markdown("---")

# =========================================================
# 11. 공통 렌더러 및 본문 실행
# =========================================================
def wrap_text_for_axis(text, max_len=14):
    text_str = str(text)
    if len(text_str) <= max_len:
        return text_str
    words = text_str.split(' ')
    lines = []
    current_line = ""
    for word in words:
        if current_line == "":
            current_line = word
        elif len(current_line) + 1 + len(word) <= max_len:
            current_line += " " + word
        else:
            lines.append(current_line)
            current_line = word
    if current_line:
        lines.append(current_line)
    return "<br>".join(lines)

def render_fullwidth_vertical_dashboard(
    title_top, 
    title_bottom, 
    table_title, 
    data_df, 
    x_col_name, 
    cat_order
):
    df = data_df.copy()
    c25_target = col_25 if col_25 in df.columns else "2025년 실적"
    c26_target = col_26 if col_26 in df.columns else "2026년 실적"
    
    if "2025년 실적" not in df.columns and c25_target in df.columns:
        df["2025년 실적"] = df[c25_target]
    if "2026년 실적" not in df.columns and c26_target in df.columns:
        df["2026년 실적"] = df[c26_target]

    if "증감액" not in df.columns:
        df["증감액"] = df["2026년 실적"] - df["2025년 실적"]
    if "증감률" not in df.columns or df["증감률"].isnull().all():
        df["증감률"] = ((df["증감액"] / df["2025년 실적"].replace(0, pd.NA)) * 100).fillna(0.0)

    display_x_col = f"{x_col_name}_wrapped"
    df[display_x_col] = df[x_col_name].apply(lambda x: wrap_text_for_axis(x, max_len=13))
    wrapped_cat_order = [wrap_text_for_axis(c, max_len=13) for c in cat_order]

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

    st.subheader(title_top)
    fig_bar = go.Figure()
    fig_bar.add_trace(go.Bar(
        x=df[display_x_col],
        y=df["2025년 실적"],
        name="2025",
        marker=dict(color="#94A3B8", line=dict(color="#64748B", width=1), cornerradius=6),
        text=label_25,
        textposition="outside",
        textfont=dict(size=14, color="#475569", family="Pretendard", weight="bold")
    ))
    fig_bar.add_trace(go.Bar(
        x=df[display_x_col],
        y=df["2026년 실적"],
        name="2026",
        marker=dict(color="#1D4ED8", line=dict(color="#1E40AF", width=1), cornerradius=6),
        text=label_26,
        textposition="outside",
        textfont=dict(size=14, color="#0F172A", family="Pretendard", weight="bold")
    ))
    fig_bar.update_layout(
        height=520,
        bargap=0.30,
        bargroupgap=0.08,
        yaxis=dict(
            rangemode='tozero',
            title=dict(text="Amount (KRW)", font=dict(size=15, color="#1E293B", weight="bold")),
            gridcolor="#F1F5F9",
            tickfont=dict(size=14, color="#475569", weight="bold")
        ),
        xaxis=dict(
            categoryorder='array',
            categoryarray=wrapped_cat_order,
            tickfont=dict(size=14, weight="bold", color="#0F172A")
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
        margin=dict(t=50, b=40, l=10, r=10)
    )
    st.plotly_chart(fig_bar, use_container_width=True)

    st.write("")
    st.markdown(f"##### {title_bottom}")
    fig_diff = go.Figure()
    fig_diff.add_trace(go.Bar(
        x=df[display_x_col],
        y=df["증감액"],
        marker=dict(color=diff_colors, cornerradius=6),
        text=diff_texts,
        textposition="outside",
        textfont=dict(size=14, family="Pretendard", weight="bold")
    ))
    fig_diff.update_layout(
        height=450,
        bargap=0.38,
        yaxis=dict(
            title=dict(text="Diff (KRW)", font=dict(size=15, color="#1E293B", weight="bold")),
            gridcolor="#F1F5F9",
            zerolinecolor="#CBD5E1",
            tickfont=dict(size=14, color="#475569", weight="bold")
        ),
        xaxis=dict(
            categoryorder='array',
            categoryarray=wrapped_cat_order,
            tickfont=dict(size=14, weight="bold", color="#0F172A")
        ),
        template="plotly_white",
        margin=dict(t=30, b=40, l=10, r=10)
    )
    st.plotly_chart(fig_diff, use_container_width=True)

    st.write("")
    st.markdown(f"##### 📋 {table_title}")
    st.dataframe(
        pd.DataFrame({
            x_col_name: df[x_col_name],
            "2025년": df["2025년 실적"],
            "2026년": df["2026년 실적"],
            "증감액": df["증감액"],
            "증감률(%)": df["증감률"]
        }),
        column_config={
            x_col_name: st.column_config.TextColumn(x_col_name, width="medium"),
            "2025년": st.column_config.NumberColumn("2025", format="₩%,d"),
            "2026년": st.column_config.NumberColumn("2026", format="₩%,d"),
            "증감액": st.column_config.NumberColumn("Diff", format="₩%+,.0f"),
            "증감률(%)": st.column_config.NumberColumn("Rate", format="%+.1f%%"),
        },
        hide_index=True,
        use_container_width=True
    )

# 페이지 분기 실행
current_page_display = t["pages"].get(page_menu, page_menu)

if page_menu == "[접수기준] 종합 실적 현황":
    st.subheader(f"🥧 {current_page_display} - Share")
    
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
            hole=0.6,
            title=t["pie_title_25"], 
            category_orders={"표준사업구분": target_categories},
            color="표준사업구분", 
            color_discrete_map=biz_colors
        )
        tot_v25 = summary_chart[col_25].sum()
        fig_pie_25.update_traces(
            textposition='inside', 
            textinfo='label+percent', 
            textfont=dict(size=14, color="#FFFFFF", weight="bold"), 
            marker=dict(line=dict(color='#FFFFFF', width=2.5))
        )
        fig_pie_25.update_layout(
            height=460, 
            title=dict(font=dict(size=17, color="#0F172A", weight="bold")),
            margin=dict(t=60, b=20, l=10, r=10), 
            legend=dict(orientation="h", yanchor="bottom", y=-0.18, xanchor="center", x=0.5),
            annotations=[dict(text=f"Total<br>{tot_v25/1e8:.1f}억", x=0.5, y=0.5, font_size=15, font_weight="bold", showarrow=False)]
        )
        st.plotly_chart(fig_pie_25, use_container_width=True)
        
    with pie_col2:
        fig_pie_26 = px.pie(
            summary_chart, 
            names="표준사업구분", 
            values=col_26, 
            hole=0.6,
            title=t["pie_title_26"], 
            category_orders={"표준사업구분": target_categories},
            color="표준사업구분", 
            color_discrete_map=biz_colors
        )
        tot_v26 = summary_chart[col_26].sum()
        fig_pie_26.update_traces(
            textposition='inside', 
            textinfo='label+percent', 
            textfont=dict(size=14, color="#FFFFFF", weight="bold"), 
            marker=dict(line=dict(color='#FFFFFF', width=2.5))
        )
        fig_pie_26.update_layout(
            height=460, 
            title=dict(font=dict(size=17, color="#0F172A", weight="bold")),
            margin=dict(t=60, b=20, l=10, r=10), 
            legend=dict(orientation="h", yanchor="bottom", y=-0.18, xanchor="center", x=0.5),
            annotations=[dict(text=f"Total<br>{tot_v26/1e8:.1f}억", x=0.5, y=0.5, font_size=15, font_weight="bold", showarrow=False)]
        )
        st.plotly_chart(fig_pie_26, use_container_width=True)

    st.write("")
    st.markdown("---")
    
    render_fullwidth_vertical_dashboard(
        title_top=f"📌 {current_page_display}",
        title_bottom="📈 Business Performance Diff & Growth Rate",
        table_title="Summary Table",
        data_df=summary_chart,
        x_col_name="표준사업구분",
        cat_order=target_categories
    )

elif page_menu == "[접수기준] 사업별 실적 현황":
    biz_filter_options = ["전체 사업 보기"] + target_categories
    current_idx = biz_filter_options.index(st.session_state.get("selected_biz_view", "전체 사업 보기"))
    selected_view = st.selectbox("Select Business:", biz_filter_options, index=current_idx, key="selected_biz_selectbox")
    if selected_view != st.session_state.get("selected_biz_view"):
        st.session_state["selected_biz_view"] = selected_view
        st.rerun()

    if selected_view == "전체 사업 보기":
        display_df = summary_chart.copy()
        x_col = "표준사업구분"
        x_categories = target_categories
    else:
        display_df = calc_summary[calc_summary["표준사업구분"] == selected_view].copy()
        x_col = "세부항목"
        x_categories = display_df["세부항목"].unique().tolist()
        
    display_df = display_df.rename(columns={col_25: "2025년 실적", col_26: "2026년 실적"})
    display_df["증감액"] = display_df["2026년 실적"] - display_df["2025년 실적"]
    display_df["증감률"] = ((display_df["증감액"] / display_df["2025년 실적"].replace(0, pd.NA)) * 100).fillna(0.0)

    render_fullwidth_vertical_dashboard(
        title_top=f"🏢 {current_page_display} ({selected_view})",
        title_bottom="📈 Detailed Performance Diff",
        table_title="Detailed Summary Table",
        data_df=display_df,
        x_col_name=x_col,
        cat_order=x_categories
    )

elif page_menu == "[접수기준] 바이어 실적 현황":
    current_tab3_biz = st.session_state.get("selected_tab3_biz", target_categories[0])
    current_idx3 = target_categories.index(current_tab3_biz) if current_tab3_biz in target_categories else 0
    selected_biz = st.selectbox("Select Business:", target_categories, index=current_idx3, key="tab3_biz_selectbox")
    if selected_biz != st.session_state.get("selected_tab3_biz"):
        st.session_state["selected_tab3_biz"] = selected_biz
        st.rerun()
    
    raw_b_chart = part_data_cache.get(selected_biz, pd.DataFrame()).copy()
    target_row = summary_chart[summary_chart["표준사업구분"] == selected_biz]
    target_tot_25 = float(target_row[col_25].values[0]) if not target_row.empty else 0.0
    target_tot_26 = float(target_row[col_26].values[0]) if not target_row.empty else 0.0

    if raw_b_chart.empty:
        st.warning("Data not found.")
    else:
        is_dash = raw_b_chart["바이어명"].astype(str).str.strip().isin(["-", "–", "—", "", "NAN", "NONE", "기타"])
        valid_buyers = raw_b_chart[~is_dash].sort_values(by="2026년 실적", ascending=False).reset_index(drop=True)
        dash_buyers = raw_b_chart[is_dash]
        
        if len(valid_buyers) > 6:
            top6 = valid_buyers.iloc[:6].copy()
            rest = valid_buyers.iloc[6:].copy()
            etc_25 = max(rest["2025년 실적"].sum() + dash_buyers["2025년 실적"].sum(), target_tot_25 - top6["2025년 실적"].sum())
            etc_26 = max(rest["2026년 실적"].sum() + dash_buyers["2026년 실적"].sum(), target_tot_26 - top6["2026년 실적"].sum())
            etc_row = pd.DataFrame([{"바이어명": "기타 (Etc)", "2025년 실적": etc_25, "2026년 실적": etc_26}])
            top_buyers = pd.concat([top6, etc_row], ignore_index=True)
        else:
            top_buyers = valid_buyers.copy()

        top_buyers["증감액"] = top_buyers["2026년 실적"] - top_buyers["2025년 실적"]
        top_buyers["증감률"] = ((top_buyers["증감액"] / top_buyers["2025년 실적"].replace(0, pd.NA)) * 100).fillna(0.0)
        x_buyer_names = [b for b in top_buyers["바이어명"] if b != "기타 (Etc)"] + (["기타 (Etc)"] if "기타 (Etc)" in top_buyers["바이어명"].values else [])

        st.subheader(f"🥧 {selected_biz} - Buyer Share")
        
        pie_col1, pie_col2 = st.columns(2)
        with pie_col1:
            fig_buyer_pie_25 = px.pie(
                top_buyers, 
                names="바이어명", 
                values="2025년 실적", 
                hole=0.6,
                title=t["buyer_pie_25"]
            )
            tot_b25 = top_buyers["2025년 실적"].sum()
            fig_buyer_pie_25.update_traces(
                textposition='inside', 
                textinfo='label+percent', 
                textfont=dict(size=14, color="#FFFFFF", weight="bold"), 
                marker=dict(line=dict(color='#FFFFFF', width=2.5))
            )
            fig_buyer_pie_25.update_layout(
                height=460, 
                title=dict(font=dict(size=17, color="#0F172A", weight="bold")),
                margin=dict(t=60, b=20, l=10, r=10), 
                legend=dict(orientation="h", yanchor="bottom", y=-0.18, xanchor="center", x=0.5),
                annotations=[dict(text=f"Total<br>{tot_b25/1e8:.1f}억" if tot_b25 >= 1e8 else f"Total<br>{tot_b25/1e4:.0f}만", x=0.5, y=0.5, font_size=14, font_weight="bold", showarrow=False)]
            )
            st.plotly_chart(fig_buyer_pie_25, use_container_width=True)
            
        with pie_col2:
            fig_buyer_pie_26 = px.pie(
                top_buyers, 
                names="바이어명", 
                values="2026년 실적", 
                hole=0.6,
                title=t["buyer_pie_26"]
            )
            tot_b26 = top_buyers["2026년 실적"].sum()
            fig_buyer_pie_26.update_traces(
                textposition='inside', 
                textinfo='label+percent', 
                textfont=dict(size=14, color="#FFFFFF", weight="bold"), 
                marker=dict(line=dict(color='#FFFFFF', width=2.5))
            )
            fig_buyer_pie_26.update_layout(
                height=460, 
                title=dict(font=dict(size=17, color="#0F172A", weight="bold")),
                margin=dict(t=60, b=20, l=10, r=10), 
                legend=dict(orientation="h", yanchor="bottom", y=-0.18, xanchor="center", x=0.5),
                annotations=[dict(text=f"Total<br>{tot_b26/1e8:.1f}억" if tot_b26 >= 1e8 else f"Total<br>{tot_b26/1e4:.0f}만", x=0.5, y=0.5, font_size=14, font_weight="bold", showarrow=False)]
            )
            st.plotly_chart(fig_buyer_pie_26, use_container_width=True)

        st.write("")
        st.markdown("---")

        render_fullwidth_vertical_dashboard(
            title_top=f"🤝 {current_page_display} ({selected_biz})",
            title_bottom="📈 Buyer Performance Diff",
            table_title="Buyer Summary Table",
            data_df=top_buyers,
            x_col_name="바이어명",
            cat_order=x_buyer_names
        )

elif page_menu == "[접수기준] 협력사 실적 현황":
    c_biz, c_vendor = st.columns([4, 6])
    with c_biz:
        selected_biz = st.selectbox("Select Business:", target_categories, key="tab4_biz_select")
    raw_v_df = vendor_data_cache.get(selected_biz, pd.DataFrame()).copy()
    if raw_v_df.empty:
        st.warning("Vendor data not found.")
    else:
        v_summary = raw_v_df.groupby("협력사명", as_index=False)[["2025년 실적", "2026년 실적"]].sum().sort_values(by="2026년 실적", ascending=False).reset_index(drop=True)
        vendor_list = ["전체 협력사 보기 (All Vendors)"] + v_summary["협력사명"].tolist()
        with c_vendor:
            selected_vendor = st.selectbox("Select Vendor:", vendor_list, key="tab4_vendor_select")
        if selected_vendor == "전체 협력사 보기 (All Vendors)":
            display_v_df = v_summary.head(6).copy()
            render_fullwidth_vertical_dashboard(
                title_top=f"🏢 {current_page_display} ({selected_biz})",
                title_bottom="📈 Vendor Performance Diff",
                table_title="Vendor Summary Table",
                data_df=display_v_df,
                x_col_name="협력사명",
                cat_order=display_v_df["협력사명"].tolist()
            )
        else:
            single_v_detail = raw_v_df[raw_v_df["협력사명"] == selected_vendor].copy()
            b_breakdown = single_v_detail.groupby("바이어명", as_index=False)[["2025년 실적", "2026년 실적"]].sum().sort_values(by="2026년 실적", ascending=False).reset_index(drop=True)
            render_fullwidth_vertical_dashboard(
                title_top=f"🏢 {selected_vendor} Performance by Buyer",
                title_bottom="📈 Breakdown Diff",
                table_title="Detailed Vendor Table",
                data_df=b_breakdown,
                x_col_name="바이어명",
                cat_order=b_breakdown["바이어명"].tolist()
            )

# =========================================================
# 12. [BI] 사업별 실적 현황 렌더링 (NameError 완벽 해결)
# =========================================================
elif page_menu.startswith("[BI_"):
    active_chart_dict = active_bi_charts.get("누계")
    if active_chart_dict is not None and not active_chart_dict.empty:
        current_bi_chart_df = active_chart_dict.copy()
    else:
        current_bi_chart_df = pd.DataFrame({
            "표준사업구분": BI_8_CATEGORIES,
            "2025년 실적": [0]*len(BI_8_CATEGORIES),
            "2026년 실적": [0]*len(BI_8_CATEGORIES),
            "증감률": [0.0]*len(BI_8_CATEGORIES)
        })
    
    if "상해" in page_menu:
        center_title_prefix = "🏭 [BI_상해]"
    elif "광주" in page_menu:
        center_title_prefix = "🏭 [BI_광주]"
    else:
        center_title_prefix = "📊 [BI_종합]"

    render_fullwidth_vertical_dashboard(
        title_top=f"{center_title_prefix} 8대 사업별 상세 실적 현황 ({display_period_name})",
        title_bottom="📈 BI 8 Categories Performance Diff",
        table_title="BI Detailed Summary Table",
        data_df=current_bi_chart_df,
        x_col_name="표준사업구분",
        cat_order=BI_8_CATEGORIES
    )
