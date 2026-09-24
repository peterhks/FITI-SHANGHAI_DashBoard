import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import io
import json
import base64
from datetime import datetime, timedelta

# =========================================================
# 1. 화면 기본 설정 및 디자인 스타일
# =========================================================
st.set_page_config(
    page_title="FITI SHANGHAI Performance Analysis",
    layout="wide",
    initial_sidebar_state="expanded"
)

MAGOK_CATEGORIES = ["법정검사", "일반검사", "섬유내수(패션잡화)", "섬유내수(중국GB)", "섬유내수(단체/정부)", "섬유수출", "연구용역", "제품인증(Q.SF)"]
OCHANG_CATEGORIES = ["산업(토목+부품)", "모빌리티(전장+의장)", "환경(환경+측정기기)", "화학바이오(화학제품+생활안전)"]
FULL_BI_CATEGORIES = MAGOK_CATEGORIES + OCHANG_CATEGORIES

# =========================================================
# 2. 사용자 권한 DB 영구 파일(JSON) 관리 시스템
# =========================================================
USER_DB_FILE = "fiti_users_db.json"
LOGIN_HISTORY_FILE = "fiti_login_history.json"

def load_user_db():
    default_db = {
        "kshan@fiti.re.kr": {"pw": "fb09010552", "role": "admin", "name": "관리자"}
    }
    if os.path.exists(USER_DB_FILE):
        try:
            with open(USER_DB_FILE, "r", encoding="utf-8") as f:
                loaded_db = json.load(f)
                if isinstance(loaded_db, dict) and len(loaded_db) > 0:
                    return loaded_db
        except Exception:
            pass
            
    with open(USER_DB_FILE, "w", encoding="utf-8") as f:
        json.dump(default_db, f, ensure_ascii=False, indent=4)
    return default_db

def save_user_db(db_data):
    with open(USER_DB_FILE, "w", encoding="utf-8") as f:
        json.dump(db_data, f, ensure_ascii=False, indent=4)

def load_login_history():
    if os.path.exists(LOGIN_HISTORY_FILE):
        try:
            with open(LOGIN_HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def save_login_history(history_data):
    with open(LOGIN_HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history_data, f, ensure_ascii=False, indent=4)

if "user_db" not in st.session_state: st.session_state["user_db"] = load_user_db()
if "login_history" not in st.session_state: st.session_state["login_history"] = load_login_history()
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
    st.session_state["current_user_email"] = ""
    st.session_state["current_user_role"] = ""
    st.session_state["current_user_name"] = ""
if "edit_target_email" not in st.session_state: st.session_state["edit_target_email"] = None

query_params = st.query_params
if "auth_ok" in query_params and query_params["auth_ok"] == "true":
    st.session_state["logged_in"] = True
    if not st.session_state["current_user_email"]:
        st.session_state["current_user_email"] = "kshan@fiti.re.kr"
        st.session_state["current_user_role"] = "admin"
        st.session_state["current_user_name"] = "관리자"

# 💡 [버그 완벽 해결] 언어 상태 충돌 방지: 최초 접속 시에만 URL 파라미터 확인, 이후에는 선택값 유지
if "lang_select" not in st.session_state:
    if "lang" in query_params:
        st.session_state["lang_select"] = "English (영어)" if query_params["lang"] == "en" else "한국어"
    else:
        st.session_state["lang_select"] = "한국어"

# =========================================================
# ⏱️ 10분 자동 로그아웃 로직 & 중국 시간(CST) 설정
# =========================================================
china_time = datetime.utcnow() + timedelta(hours=8)

if st.session_state["logged_in"]:
    if "last_active_time" in st.session_state:
        if (china_time - st.session_state["last_active_time"]).total_seconds() > 600:
            st.session_state["logged_in"] = False
            st.session_state["current_user_email"] = ""
            st.session_state["current_user_role"] = ""
            st.session_state["current_user_name"] = ""
            if "auth_ok" in st.query_params: del st.query_params["auth_ok"]
            del st.session_state["last_active_time"]
            st.session_state["auto_logout_alert"] = True
            st.rerun()
    st.session_state["last_active_time"] = china_time

# =========================================================
# 3. 다국어 텍스트 사전 및 언어 설정
# =========================================================
LANG_DICT = {
    "한국어": {
        "sys_title": "상해지사 실적 종합 분석 시스템",
        "sys_sub": "飞迪商品检验（上海）有限公司 | 상해지사 사업팀",
        "data_mgmt": "📁 데이터 관리",
        "admin_upload": "공용 엑셀 업로드 (관리자 전용)",
        "admin_caption": "💡 엑셀 교체는 관리자 권한이 필요합니다.",
        "sync_success": "✅ 서버 파일 및 세션 동기화 완료!",
        "shared_file_info": "📂 서버 공용 최신 파일 연동 중",
        "file_not_found": "분석할 엑셀 파일을 찾을 수 없습니다. 관리자 계정으로 업로드해 주세요.",
        "page_select": "📑 분석 페이지 선택",
        "cat_select": "📌 카테고리 선택",
        "period_select": "⏱️ [BI] 실적 기간 선택",
        "kpi_25": "📅 25년 총 실적",
        "kpi_26": "🚀 26년 총 실적",
        "kpi_diff": "📈 실적 증감액",
        "kpi_rate": "📊 증감 퍼센트",
        "kpi_diff_sub": "전년 대비 실적차",
        "kpi_rate_sub": "전년 대비 성장률",
        "pie_title_25": "2025년 사업별 실적 비중",
        "pie_title_26": "2026년 사업별 실적 비중",
        "unit": "원",
        "welcome": "님 환영합니다.",
        "role_label": "권한",
        "role_admin": "관리자",
        "role_bi": "접수 + BI",
        "role_gen": "접수",
        "account_info": "접속 계정",
        "logout": "시스템 잠금 (로그아웃)",
        "admin_menu": "🛡️ 관리자 권한",
        "bi_tot": "BI_종합",
        "bi_sh": "BI_상해",
        "bi_gw": "BI_광주",
        "center_compare_title": "마곡 본원 vs 오창 분원 거점별 실적 비교",
        "center_pie_25": "2025년 거점별 실적 비중",
        "center_pie_26": "2026년 거점별 실적 비중",
        "center_growth": "마곡 본원 vs 오창 분원 요약 비교",
        "magok_title": "마곡 본원 세부 사업별 실적 현황",
        "ochang_title": "오창 분원 세부 사업별 실적 현황",
        "all_cat_title": "전체 12대 사업별 상세 실적 현황",
        "all_biz": "전체 사업 보기",
        "all_vendor": "전체 협력사 보기",
        "perf_status": " 실적 현황",
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
    "English (영어)": {
        "sys_title": "Shanghai Branch Performance Analysis System",
        "sys_sub": "FITI Testing & Research Institute Shanghai | Business Team",
        "data_mgmt": "📁 Data Management",
        "admin_upload": "Upload Public Excel (Admin Only)",
        "admin_caption": "💡 Admin permission required for Excel replacement.",
        "sync_success": "✅ Server public file & session synced!",
        "shared_file_info": "📂 Syncing latest public file",
        "file_not_found": "Excel file not found. Please login as admin to upload.",
        "page_select": "📑 Select Page",
        "cat_select": "📌 Select Category",
        "period_select": "⏱️ [BI] Period Select",
        "kpi_25": "📅 '25 Total Performance",
        "kpi_26": "🚀 '26 Total Performance",
        "kpi_diff": "📈 Performance Diff",
        "kpi_rate": "📊 Growth Rate",
        "kpi_diff_sub": "YoY Performance Gap",
        "kpi_rate_sub": "YoY Growth Rate",
        "pie_title_25": "2025 Performance Share by Business",
        "pie_title_26": "2026 Performance Share by Business",
        "unit": "KRW",
        "welcome": "Welcome,",
        "role_label": "Role",
        "role_admin": "Admin",
        "role_bi": "Reception + BI",
        "role_gen": "Reception",
        "account_info": "Account",
        "logout": "Lock System (Logout)",
        "admin_menu": "🛡️ Admin Privileges",
        "bi_tot": "BI_Total",
        "bi_sh": "BI_Shanghai",
        "bi_gw": "BI_Gwangju",
        "center_compare_title": "Magok vs Ochang Performance Comparison",
        "center_pie_25": "2025 Share by Center",
        "center_pie_26": "2026 Share by Center",
        "center_growth": "Magok vs Ochang Summary",
        "magok_title": "Magok Detailed Performance",
        "ochang_title": "Ochang Detailed Performance",
        "all_cat_title": "All 12 Categories Detailed Performance",
        "all_biz": "All Businesses",
        "all_vendor": "All Vendors",
        "perf_status": " Performance",
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

t = LANG_DICT.get(st.session_state["lang_select"], LANG_DICT["한국어"])

# =========================================================
# 4. 스타일 및 디자인
# =========================================================
st.markdown("""
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable.min.css');
    html, body, [class*="css"] {
        font-family: 'Pretendard Variable', Pretendard, -apple-system, BlinkMacSystemFont, 'Apple SD Gothic Neo', system-ui, Roboto, sans-serif !important;
    }
    section[data-testid="stSidebar"] { padding-top: 0rem !important; }
    section[data-testid="stSidebar"] div.block-container { padding-top: 0.1rem !important; padding-bottom: 0.4rem !important; }
    section[data-testid="stSidebar"] div.stExpander { margin-bottom: 0.1rem !important; }
    section[data-testid="stSidebar"] hr { margin: 0.2rem 0 !important; }
    
    div[data-testid="stSidebar"] div.stExpander button[kind="secondary"] {
        background-color: transparent !important; border: none !important; color: #334155 !important; padding: 0 !important; box-shadow: none !important; font-size: 13px !important; justify-content: flex-start !important; height: auto !important; min-height: 0 !important; margin-top: -2px !important; margin-bottom: -2px !important;
    }
    div[data-testid="stSidebar"] div.stExpander button[kind="secondary"]:hover { color: #2563EB !important; text-decoration: underline; }

    .fiti-header {
        background: linear-gradient(135deg, #002B5C 0%, #003876 100%);
        padding: 22px 28px; border-radius: 10px; display: flex; align-items: center; gap: 22px; color: #FFFFFF; margin-bottom: 22px; box-shadow: 0 4px 14px rgba(0, 43, 92, 0.18);
    }
    .fiti-logo-text { font-size: 28px; font-weight: 900; letter-spacing: -0.5px; border-right: 1.5px solid rgba(255, 255, 255, 0.25); padding-right: 22px; }
    .fiti-title-main { font-size: 21px; font-weight: 800; margin-bottom: 4px; letter-spacing: -0.3px; }
    .fiti-title-sub { font-size: 13px; color: #D0E1FD; font-weight: 400; }
    .kpi-card {
        background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px; padding: 20px 22px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04); border-top: 4px solid #CBD5E1;
    }
    .kpi-title { font-size: 13px; font-weight: 600; color: #64748B; margin-bottom: 8px; }
    .kpi-num { font-size: 26px; font-weight: 800; color: #0F172A; letter-spacing: -0.5px; }
    .kpi-sub { font-size: 12px; color: #94A3B8; margin-top: 6px; }
    .kpi-badge { display: inline-block; padding: 3px 8px; border-radius: 6px; font-size: 12px; font-weight: 700; margin-top: 6px; }
    .sidebar-card-btn, .sidebar-card-btn-active {
        display: block; width: 100%; border-radius: 8px; text-align: center; font-size: 13px; padding: 6px 8px; margin-bottom: 2px; text-decoration: none;
    }
</style>
""", unsafe_allow_html=True)

# =========================================================
# 5. 상해 야경 테마 로그인 화면
# =========================================================
if not st.session_state["logged_in"]:
    if st.session_state.get("auto_logout_alert"):
        st.warning("⏱️ 10분 동안 활동이 없어 보안을 위해 자동으로 로그아웃되었습니다.")
        st.session_state["auto_logout_alert"] = False

    bg_image_path = "fiti_shanghai_bg.png"
    if os.path.exists(bg_image_path):
        with open(bg_image_path, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
        bg_css = f"background: linear-gradient(rgba(0, 15, 35, 0.4), rgba(0, 10, 25, 0.55)), url('data:image/png;base64,{encoded_string}') no-repeat center center fixed; background-size: cover;"
    else:
        bg_css = "background: linear-gradient(135deg, #001E3D 0%, #000B1A 100%);"

    st.markdown(f"""
    <style>
        .stMain {{ {bg_css} min-height: 100vh; }}
        header {{visibility: hidden;}}
        [data-testid="stForm"] {{ max-width: 420px !important; margin: 6vh auto !important; background: transparent !important; border: none !important; box-shadow: none !important; padding: 10px !important; }}
        [data-testid="stForm"] [data-testid="stTextInput"] div[data-baseweb="input"] {{ width: 100% !important; background: rgba(255, 255, 255, 0.9) !important; border-radius: 8px !important; }}
        [data-testid="stForm"] .stFormSubmitButton button {{ width: 100% !important; font-weight: 800 !important; font-size: 15px !important; background-color: #3B82F6 !important; color: #FFFFFF !important; border: none !important; border-radius: 8px !important; padding: 10px !important; }}
    </style>
    """, unsafe_allow_html=True)
    
    with st.form("login_form"):
        st.markdown("""
        <div style="text-align: center; margin-bottom: 25px; color: #FFFFFF;">
            <div style="font-size: 48px; font-weight: 900; margin-bottom: 6px;">FITI Shanghai</div>
            <div style="font-size: 18px; font-weight: 800; margin-bottom: 4px;">상해지사 실적 종합 분석 시스템</div>
            <div style="font-size: 13px; color: #38BDF8; font-weight: 600;">飞迪商品检验（上海）有限公司 | 상해지사 사업팀</div>
        </div>
        """, unsafe_allow_html=True)
        
        login_input_raw = st.text_input("아이디 또는 이메일", placeholder="예: gdhong", label_visibility="collapsed")
        login_pw = st.text_input("비밀번호", type="password", placeholder="비밀번호 입력", label_visibility="collapsed")
        if st.form_submit_button("시스템 로그인", use_container_width=True):
            raw_val = login_input_raw.strip()
            candidate_emails = [raw_val]
            if raw_val and "@" not in raw_val:
                candidate_emails.extend([f"{raw_val}@fiti.re.kr", f"{raw_val}@fitiglobal.com"])
            
            matched_email, user_record = None, None
            for ce in candidate_emails:
                if ce in st.session_state["user_db"]:
                    matched_email = ce
                    user_record = st.session_state["user_db"][ce]
                    break
            
            if user_record and user_record["pw"] == login_pw.strip():
                st.session_state["logged_in"] = True
                st.session_state["current_user_email"] = matched_email
                st.session_state["current_user_role"] = user_record["role"]
                st.session_state["current_user_name"] = user_record["name"]
                st.session_state["login_history"].append({"time": china_time.strftime("%Y-%m-%d %H:%M:%S"), "email": matched_email, "name": user_record["name"], "status": "성공"})
                save_login_history(st.session_state["login_history"])
                st.query_params["auth_ok"] = "true"
                st.rerun()
            else:
                st.session_state["login_history"].append({"time": china_time.strftime("%Y-%m-%d %H:%M:%S"), "email": raw_val if raw_val else "입력없음", "name": "미인증", "status": "실패"})
                save_login_history(st.session_state["login_history"])
                st.error("아이디 또는 비밀번호가 일치하지 않습니다.")
    st.stop()

# =========================================================
# 6. 상단 공식 배너
# =========================================================
welcome_str = f"Welcome, <b>{st.session_state['current_user_name']}</b>" if st.session_state["lang_select"] == "English (영어)" else f"<b>{st.session_state['current_user_name']}</b>{t['welcome']}"
role_str = t['role_admin'] if st.session_state['current_user_role']=='admin' else (t['role_bi'] if st.session_state['current_user_role']=='bi_user' else t['role_gen'])

st.markdown(f"""
<div class="fiti-header">
    <div class="fiti-logo-text">FITI</div>
    <div style="flex-grow: 1;">
        <div class="fiti-title-main">{t["sys_title"]}</div>
        <div class="fiti-title-sub">{t["sys_sub"]}</div>
    </div>
    <div style="text-align: right; font-size: 13px; color: #D0E1FD;">
        {welcome_str}<br>
        <span style="background: rgba(255,255,255,0.2); padding: 2px 8px; border-radius: 4px; font-size: 11px;">
            {t['role_label']}: {role_str}
        </span>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================
# 7. 파일 로드
# =========================================================
EXCEL_FILE = "performance_최신.xlsx"
os.makedirs("downloads", exist_ok=True)
LOCAL_EXCEL_PATH = os.path.join("downloads", EXCEL_FILE)

if "persistent_file_bytes" in st.session_state:
    raw_bytes = st.session_state["persistent_file_bytes"]
elif os.path.exists(LOCAL_EXCEL_PATH):
    with open(LOCAL_EXCEL_PATH, "rb") as f: raw_bytes = f.read()
    st.session_state["persistent_file_bytes"] = raw_bytes
else:
    raw_bytes = None

try:
    temp_stream = io.BytesIO(raw_bytes)
    total_sheets_count = len(pd.ExcelFile(temp_stream).sheet_names)
except Exception:
    total_sheets_count = 0

# =========================================================
# 8. 사이드바 네비게이션
# =========================================================
st.sidebar.selectbox("🌐 Language", ["한국어", "English (영어)"], key="lang_select", label_visibility="collapsed")
current_lang_code = "en" if st.session_state["lang_select"] == "English (영어)" else "ko"

# 💡 [버그 완벽 해결] 사이드바 언어 갱신 시 주소창(query param)에도 언어 코드 실시간 적용
st.query_params["lang"] = current_lang_code

user_role = st.session_state["current_user_role"]
all_pages_keys = ["[접수기준] 종합 실적 현황", "[접수기준] 사업별 실적 현황", "[접수기준] 바이어 실적 현황", "[접수기준] 협력사 실적 현황"] if user_role == "general_user" else [
    "[접수기준] 종합 실적 현황", "[접수기준] 사업별 실적 현황", "[접수기준] 바이어 실적 현황", "[접수기준] 협력사 실적 현황",
    "[BI_종합] 사업별 실적 현황", "[BI_상해] 사업별 실적 현황", "[BI_광주] 사업별 실적 현황"
]

st.sidebar.markdown(f"#### {t['page_select']}")
if "current_page" not in st.session_state or st.session_state["current_page"] not in all_pages_keys: st.session_state["current_page"] = all_pages_keys[0]
if "page" in query_params and query_params["page"] in all_pages_keys: st.session_state["current_page"] = query_params["page"]

for p_key in all_pages_keys:
    is_active = (st.session_state["current_page"] == p_key)
    btn_class = "sidebar-card-btn-active" if is_active else "sidebar-card-btn"
    display_name = t["pages"].get(p_key, p_key)
    # 카테고리 링크 클릭 시에도 &lang={current_lang_code} 가 붙어서 리셋을 완벽 방어함
    st.sidebar.markdown(f'<a href="?page={p_key}&auth_ok=true&lang={current_lang_code}" class="{btn_class}" style="font-weight: {"800" if is_active else "700"}; background-color: {"#003876" if is_active else "#F8FAFC"}; color: {"#FFFFFF" if is_active else "#0F172A"}; border: 1.5px solid {"#001E3D" if is_active else "#CBD5E1"};" target="_self">{display_name}</a>', unsafe_allow_html=True)

page_menu = st.session_state["current_page"]
kpi_container = st.sidebar.container()

# =========================================================
# 11. 사이드바 하단: [접속 계정] 및 [관리자 권한]
# =========================================================
st.sidebar.markdown("---")
st.sidebar.markdown(f"👤 **{t['account_info']}**: {st.session_state['current_user_name']}")
if st.sidebar.button(t['logout'], use_container_width=True):
    for k in ["logged_in", "current_user_email", "current_user_role", "current_user_name"]: st.session_state[k] = ""
    st.session_state["logged_in"] = False
    if "auth_ok" in st.query_params: del st.query_params["auth_ok"]
    st.rerun()

if user_role in ["admin", "bi_user"]:
    with st.sidebar.expander(t["admin_menu"], expanded=False):
        st.markdown(f"##### {t['data_mgmt']}")
        if st.session_state["current_user_role"] == "admin":
            uploaded_file = st.file_uploader(t["admin_upload"], type=["xlsx", "csv"], label_visibility="collapsed")
            if uploaded_file is not None:
                file_bytes = uploaded_file.getvalue()
                with open(LOCAL_EXCEL_PATH, "wb") as f: f.write(file_bytes)
                st.session_state["persistent_file_bytes"] = file_bytes
                st.cache_data.clear()
                st.success(t["sync_success"])
                st.rerun()
        else:
            st.caption(t["admin_caption"])
            
        st.markdown(f"<p style='font-size:10px; color:#64748B; margin-top:-5px; margin-bottom:2px; white-space:nowrap;'>📂 파일 연동 중 (시트수: {total_sheets_count}개)</p>", unsafe_allow_html=True)
            
        st.markdown("---")
        st.markdown("#### 👥 등록된 담당자 목록")
        
        cat_reception = {em: info for em, info in st.session_state["user_db"].items() if info["role"] == "general_user"}
        cat_reception_bi = {em: info for em, info in st.session_state["user_db"].items() if info["role"] == "bi_user"}
        cat_admin = {em: info for em, info in st.session_state["user_db"].items() if info["role"] == "admin"}
        
        st.markdown("<div style='margin-top: -5px; margin-bottom: 2px;'><b>[ 접수 ]</b></div>", unsafe_allow_html=True)
        if cat_reception:
            for em, info in cat_reception.items():
                if st.button(f"• {info['name']} ({em})", key=f"btn_{em}", use_container_width=True):
                    st.session_state["edit_target_email"] = em
                    st.rerun()
        else:
            st.markdown("<div style='font-size:11px; margin-top: -5px; padding-left: 10px; color:gray;'>• 인원 없음</div>", unsafe_allow_html=True)
            
        st.markdown("<div style='margin-top: 5px; margin-bottom: 2px;'><b>[ 접수 + BI ]</b></div>", unsafe_allow_html=True)
        if cat_reception_bi:
            for em, info in cat_reception_bi.items():
                if st.button(f"• {info['name']} ({em})", key=f"btn_{em}", use_container_width=True):
                    st.session_state["edit_target_email"] = em
                    st.rerun()
        else:
            st.markdown("<div style='font-size:11px; margin-top: -5px; padding-left: 10px; color:gray;'>• 인원 없음</div>", unsafe_allow_html=True)
            
        st.markdown("<div style='margin-top: 5px; margin-bottom: 2px;'><b>[ 관리자 ]</b></div>", unsafe_allow_html=True)
        if cat_admin:
            for em, info in cat_admin.items():
                if st.button(f"• {info['name']} ({em})", key=f"btn_{em}", use_container_width=True):
                    st.session_state["edit_target_email"] = em
                    st.rerun()
        else:
            st.markdown("<div style='font-size:11px; margin-top: -5px; padding-left: 10px; color:gray;'>• 인원 없음</div>", unsafe_allow_html=True)
            
        st.markdown("---")
        
        edit_email = st.session_state.get("edit_target_email", None)
        is_editing = edit_email is not None and edit_email in st.session_state["user_db"]
        
        if is_editing:
            st.markdown(f"#### ✏️ 담당자 수정 ({edit_email})")
            curr_info = st.session_state["user_db"][edit_email]
            default_name, default_pw, curr_role = curr_info["name"], curr_info["pw"], curr_info["role"]
            default_cat_idx = 0 if curr_role == "general_user" else (1 if curr_role == "bi_user" else 2)
        else:
            st.markdown("#### ➕ 담당자 추가")
            default_name, default_pw, default_cat_idx = "", "", 0

        with st.form("add_user_form"):
            new_email = st.text_input("아이디 또는 이메일", value=edit_email if is_editing else "", placeholder="예: gdhong", disabled=is_editing, label_visibility="collapsed")
            new_name = st.text_input("담당자 성명", value=default_name, placeholder="홍길동", label_visibility="collapsed")
            new_pw = st.text_input("비밀번호", value=default_pw, type="password", placeholder="비밀번호 입력", label_visibility="collapsed")
            new_category = st.selectbox("권한", ["접수", "접수 + BI", "관리자"], index=default_cat_idx, label_visibility="collapsed")
            
            btn_label = "담당자 정보 수정" if is_editing else "담당자 등록"
            if st.form_submit_button(btn_label, type="primary", use_container_width=True):
                target_key = edit_email if is_editing else new_email.strip()
                if target_key and new_pw.strip():
                    if "@" not in target_key: target_key = f"{target_key}@fiti.re.kr"
                    assigned_role = "general_user" if new_category == "접수" else ("bi_user" if new_category == "접수 + BI" else "admin")
                    st.session_state["user_db"][target_key] = {"pw": new_pw.strip(), "role": assigned_role, "name": new_name.strip() if new_name.strip() else target_key}
                    save_user_db(st.session_state["user_db"])
                    st.session_state["edit_target_email"] = None
                    st.success(f"✅ {'수정' if is_editing else '등록'} 완료!")
                    st.rerun()
                else: st.error("아이디와 비밀번호는 필수입니다.")
        
        if is_editing:
            if st.button("➕ 신규 등록 모드로 전환", type="primary", use_container_width=True):
                st.session_state["edit_target_email"] = None
                st.rerun()

        target_delete_email = st.selectbox("삭제할 담당자 선택", ["선택하세요."] + list(st.session_state["user_db"].keys()), label_visibility="collapsed")
        if st.button("선택한 담당자 삭제", type="primary", use_container_width=True):
            if target_delete_email != "선택하세요.":
                if target_delete_email == st.session_state["current_user_email"]:
                    st.error("현재 로그인 중인 계정은 삭제 불가합니다.")
                else:
                    del st.session_state["user_db"][target_delete_email]
                    save_user_db(st.session_state["user_db"])
                    if st.session_state.get("edit_target_email") == target_delete_email: st.session_state["edit_target_email"] = None
                    st.success("🗑️ 영구 삭제 완료!")
                    st.rerun()

    with st.sidebar.expander("📋 최근 로그인 감사 로그", expanded=False):
        if st.session_state["login_history"]:
            df_log = pd.DataFrame(st.session_state["login_history"])
            df_log["날짜"] = pd.to_datetime(df_log["time"]).dt.date
            unique_dates = ["전체 날짜 보기"] + sorted(df_log["날짜"].astype(str).unique().tolist(), reverse=True)
            sel_date = st.selectbox("날짜별 로그", unique_dates, label_visibility="collapsed")
            f_log_df = df_log[df_log["날짜"].astype(str) == sel_date] if sel_date != "전체 날짜 보기" else df_log
            st.dataframe(f_log_df[["time", "email", "name", "status"]].tail(10), hide_index=True, use_container_width=True)
            st.download_button("📥 다운로드 (CSV)", data=f_log_df.to_csv(index=False).encode('utf-8-sig'), file_name=f"fiti_log_{china_time.strftime('%Y%m%d')}.csv", mime="text/csv", use_container_width=True)
        else:
            st.caption("기록 없음")

if not raw_bytes:
    st.warning(t["file_not_found"])
    st.stop()

# =========================================================
# 10. 초고속 데이터 파싱
# =========================================================
@st.cache_data
def load_and_parse_all_data(file_bytes_val):
    def clean_series(series):
        cleaned = series.astype(str).str.replace(',', '').str.replace('₩', '').str.strip()
        cleaned = cleaned.replace(['-', '–', '—', 'nan', 'NaN', 'None', ''], '0')
        return pd.to_numeric(cleaned, errors='coerce').fillna(0)

    stream = io.BytesIO(file_bytes_val)
    excel_obj = pd.ExcelFile(stream)
    all_sheets = excel_obj.sheet_names
    total_sheets_count = len(all_sheets)
    sheet_dict = {s.strip().lower().replace(" ", "").replace("_", ""): s for s in all_sheets}

    def get_sheet(keywords):
        for s_clean, orig_name in sheet_dict.items():
            if all(k.lower().replace(" ", "").replace("_", "") in s_clean for k in keywords): return orig_name
        return None

    # --- 1) 종합 요약 파싱 ---
    summary_sheet_name = get_sheet(["종합"]) or all_sheets[0]
    raw_summary = pd.read_excel(stream, sheet_name=summary_sheet_name, header=None)
    h_idx = next((idx for idx, row in raw_summary.iterrows() if "구분" in "".join(row.dropna().astype(str).tolist()) and any(k in "".join(row.dropna().astype(str).tolist()) for k in ["합계", "25", "26"])), 0)
    
    stream.seek(0)
    df_summary = pd.read_excel(stream, sheet_name=summary_sheet_name, skiprows=h_idx)
    for c in df_summary.columns:
        if df_summary[c].dtype == object:
            conv = pd.to_numeric(df_summary[c].astype(str).str.replace(',', '').str.strip(), errors='coerce')
            if conv.notnull().mean() > 0.5: df_summary[c] = conv.fillna(0)

    num_cols = df_summary.select_dtypes(include=['number']).columns.tolist()
    col_25 = next((c for c in num_cols if "25" in str(c)), num_cols[0] if num_cols else None)
    col_26 = next((c for c in num_cols if "26" in str(c)), num_cols[1] if len(num_cols) > 1 else num_cols[0])
    
    other_cols = [c for c in df_summary.columns if c not in num_cols]
    cat_col = next((c for c in other_cols if any(k in "".join(df_summary[c].dropna().astype(str).tolist()) for k in ["패션잡화", "GB", "글로벌", "제품평가"])), other_cols[0] if other_cols else df_summary.columns[0])
    sub_cat_col = other_cols[1] if len(other_cols) > 1 else None

    df_summary[cat_col] = df_summary[cat_col].replace(r'^\s*$', pd.NA, regex=True)
    df_summary["사업구분_채움"] = df_summary[cat_col].ffill()

    def map_biz_category(val):
        s = str(val).replace(" ", "").upper()
        if "글로벌" in s or "GLOBAL" in s: return "글로벌 바이어"
        elif "패션" in s or "잡화" in s or "KC" in s: return "패션잡화"
        elif "GB" in s: return "GB"
        elif "제품평가" in s or "INSPECTION" in s: return "제품평가"
        return None

    df_summary["표준사업구분"] = df_summary["사업구분_채움"].apply(map_biz_category)
    calc_summary = df_summary[(~df_summary[cat_col].astype(str).str.strip().str.upper().str.contains(r"SUB\s*TOTAL|TOTAL|합계|소계", regex=True, na=False)) & (df_summary["표준사업구분"].notnull())].copy()
    calc_summary["세부항목"] = calc_summary[sub_cat_col].fillna(calc_summary["표준사업구분"]).astype(str) if sub_cat_col else calc_summary["표준사업구분"]
    
    target_cats = ["글로벌 바이어", "패션잡화", "GB", "제품평가"]
    summary_chart = calc_summary.groupby("표준사업구분", as_index=False)[[col_25, col_26]].sum()
    summary_chart["정렬"] = summary_chart["표준사업구분"].apply(lambda x: target_cats.index(x) if x in target_cats else 99)
    summary_chart = summary_chart.sort_values("정렬").reset_index(drop=True)

    # --- 2) 바이어 및 협력사 파싱 ---
    PART_MAP = {"패션잡화": [["kc"]], "GB": [["gb"]], "글로벌 바이어": [["global", "1"], ["global", "2"]], "제품평가": [["inspection", "원단"], ["inspection", "가먼트"]]}
    part_cache, vendor_cache = {}, {}
    
    for cat in target_cats:
        b_dfs, v_dfs = [], []
        for kw in PART_MAP.get(cat, []):
            sh_name = get_sheet(kw)
            if not sh_name: continue
            
            stream.seek(0)
            raw = pd.read_excel(stream, sheet_name=sh_name, header=None)
            
            pr, pc = next(((r, c) for r in range(min(20, len(raw))) for c in range(len(raw.columns)) if "행레이블" in str(raw.iat[r, c]).replace(" ", "")), (None, None))
            if pr is not None:
                sub_r = raw.iloc[pr:, pc:pc+6].copy().reset_index(drop=True)
                sub_r.columns = [str(c).strip() for c in sub_r.iloc[0]]
                sub_d = sub_r.iloc[1:].copy().reset_index(drop=True)
                bc = sub_d.columns[0]
                c25_b = next((c for c in sub_d.columns[1:] if "25" in str(c).replace(" ", "") and ("합계" in str(c) or "실적" in str(c))), None)
                c26_b = next((c for c in sub_d.columns[1:] if "26" in str(c).replace(" ", "") and ("합계" in str(c) or "실적" in str(c))), None)
                if c25_b and c26_b:
                    p_rows = []
                    for _, row in sub_d.iterrows():
                        b_name = str(row[bc]).strip()
                        if not b_name or b_name.lower() in ['nan', 'none']: continue
                        if any(k in b_name.replace(" ", "") for k in ["총합계", "합계", "전체합계"]): break
                        p_rows.append({"바이어명": b_name, "2025년 실적": clean_series(pd.Series([row[c25_b]])).iloc[0], "2026년 실적": clean_series(pd.Series([row[c26_b]])).iloc[0]})
                    if p_rows: b_dfs.append(pd.DataFrame(p_rows))
            
            hr, b_c_idx, v_c_idx = None, None, None
            for r in range(min(15, len(raw))):
                r_vals = [str(x).strip().replace(" ", "") for x in raw.iloc[r].tolist()]
                for c, val in enumerate(r_vals):
                    if "업체명" in val or "협력사" in val: v_c_idx = c
                    elif "바이어" in val: b_c_idx = c
                if v_c_idx is not None: hr = r; break
            if hr is not None and v_c_idx is not None:
                stream.seek(0)
                df_r = pd.read_excel(stream, sheet_name=sh_name, skiprows=hr)
                v_c, b_c = df_r.columns[v_c_idx], df_r.columns[b_c_idx] if b_c_idx is not None and b_c_idx < len(df_r.columns) else None
                c25_v = next((c for c in df_r.columns if "25" in str(c).replace(" ", "") and ("합계" in str(c) or "실적" in str(c))), None)
                c26_v = next((c for c in df_r.columns if "26" in str(c).replace(" ", "") and ("합계" in str(c) or "실적" in str(c))), None)
                if c25_v and c26_v:
                    df_c = df_r.dropna(subset=[v_c]).copy()
                    df_c = df_c[~df_c[v_c].astype(str).str.contains(r"소계|합계|TOTAL|총계", regex=True, na=False)]
                    res = pd.DataFrame()
                    res["협력사명"] = df_c[v_c].astype(str).str.strip()
                    res["바이어명"] = df_c[b_c].astype(str).str.strip() if b_c else "기본"
                    res["2025년 실적"], res["2026년 실적"] = clean_series(df_c[c25_v]), clean_series(df_c[c26_v])
                    v_dfs.append(res[res["협력사명"] != ""])

        part_cache[cat] = pd.concat(b_dfs, ignore_index=True).groupby("바이어명", as_index=False)[["2025년 실적", "2026년 실적"]].sum() if b_dfs else pd.DataFrame(columns=["바이어명", "2025년 실적", "2026년 실적"])
        vendor_cache[cat] = pd.concat(v_dfs, ignore_index=True) if v_dfs else pd.DataFrame(columns=["협력사명", "바이어명", "2025년 실적", "2026년 실적"])

    # --- 3) BI 분석 파싱 ---
    def parse_bi(branch_name):
        s_k = branch_name.strip().lower().replace(" ", "").replace("_", "")
        t_sn = None
        for s_orig in all_sheets:
            s_clean = s_orig.strip().lower().replace(" ", "").replace("_", "")
            if s_k == "광주" and s_clean == "bi광주": t_sn = s_orig; break
            elif s_k == "상해" and s_clean == "bi상해": t_sn = s_orig; break
            elif s_k == "종합" and s_clean in ["bi종합", "bi"]: t_sn = s_orig; break
            
        if not t_sn:
            for s_orig in all_sheets:
                s_clean = s_orig.strip().lower().replace("_", "").replace(" ", "")
                if s_k == "광주" and "광주" in s_clean and "상해" not in s_clean: t_sn = s_orig; break
                elif s_k == "상해" and "상해" in s_clean and "광주" not in s_clean: t_sn = s_orig; break
                elif s_k == "종합" and "bi" in s_clean and "광주" not in s_clean and "상해" not in s_clean: t_sn = s_orig; break

        emp_kpi = {"25": 0, "26": 0, "diff": 0, "rate": 0.0}
        emp_df = pd.DataFrame({"표준사업구분": FULL_BI_CATEGORIES, "2025년 실적": [0]*len(FULL_BI_CATEGORIES), "2026년 실적": [0]*len(FULL_BI_CATEGORIES), "증감률": [0.0]*len(FULL_BI_CATEGORIES)})
        if not t_sn: return {"전체 총계 누계": emp_kpi, "사업 소계 누계": emp_kpi, "사업 소계 월계": emp_kpi}, {"누계": emp_df, "월계": emp_df}
        
        stream.seek(0)
        raw = pd.read_excel(stream, sheet_name=t_sn, header=None)
        m_25, m_26, m_r, c_25, c_26, c_r = None, None, None, None, None, None
        for r in range(min(20, len(raw))):
            for c in range(len(raw.columns)):
                v = str(raw.iat[r, c]).strip().replace(" ", "")
                if v == "월계" and m_25 is None: m_25, m_26, m_r = c, c+1, c+2
                elif v == "누계" and c_25 is None: c_25, c_26, c_r = c, c+1, c+2
        
        tr_idx, sr_idx = None, None
        for idx in range(len(raw)):
            r_str = "".join(raw.iloc[idx].dropna().astype(str).tolist()).replace(" ", "")
            if "전체총계" in r_str: tr_idx = idx
            elif "사업소계" in r_str and sr_idx is None: sr_idx = idx

        def ev(ri, c25i, c26i, ratei):
            if ri is None or c25i is None or c25i >= len(raw.columns): return emp_kpi
            try:
                r = raw.iloc[ri]
                v25, v26 = clean_series(pd.Series([r.iat[c25i]])).iloc[0]*1000, clean_series(pd.Series([r.iat[c26i]])).iloc[0]*1000
                rt = clean_series(pd.Series([r.iat[ratei]])).iloc[0] if ratei < len(raw.columns) else 0.0
                return {"25": v25, "26": v26, "diff": v26 - v25, "rate": rt if rt != 0 else (round((v26 - v25)/v25*100, 1) if v25 != 0 else 0.0)}
            except: return emp_kpi

        t_map = [
            ("법정검사", ["법정검사", "법정"], "합계"), ("일반검사", ["일반검사", "일반"], "합계"), ("섬유내수(패션잡화)", ["패션잡화", "패션"], "소계"),
            ("섬유내수(중국GB)", ["중국gb", "gb"], "소계"), ("섬유내수(단체/정부)", ["단체/정부", "단체", "정부"], "소계"), ("섬유수출", ["섬유수출", "수출"], "합계"),
            ("연구용역", ["연구용역", "연구"], "합계"), ("제품인증(Q.SF)", ["제품인증", "q.sf", "sf"], "합계"), ("산업(토목+부품)", ["산업", "토목", "부품"], "합계"),
            ("모빌리티(전장+의장)", ["모빌리티", "전장", "의장"], "합계"), ("환경(환경+측정기기)", ["환경", "측정"], "합계"), ("화학바이오(화학제품+생활안전)", ["화학", "바이오", "생활안전"], "합계")
        ]

        def bc(c25i, c26i, ratei):
            res = []
            for cat, kws, mt in t_map:
                midx = None
                if c25i is not None:
                    for idx in range(len(raw)):
                        txt = " ".join([str(raw.iat[idx, c]) for c in range(len(raw.columns))]).replace(" ", "").lower()
                        if any(kw.lower().replace(" ", "") in txt for kw in kws) and mt in txt: midx = idx; break
                if midx is not None:
                    r = raw.iloc[midx]
                    v25, v26 = clean_series(pd.Series([r.iat[c25i]])).iloc[0]*1000, clean_series(pd.Series([r.iat[c26i]])).iloc[0]*1000
                    rt = clean_series(pd.Series([r.iat[ratei]])).iloc[0] if ratei < len(raw.columns) else 0.0
                    if rt == 0.0 and v25 != 0: rt = round((v26 - v25)/v25*100, 1)
                    res.append({"표준사업구분": cat, "2025년 실적": v25, "2026년 실적": v26, "증감률": rt})
                else: res.append({"표준사업구분": cat, "2025년 실적": 0, "2026년 실적": 0, "증감률": 0.0})
            return pd.DataFrame(res)

        return {"전체 총계 누계": ev(tr_idx, c_25, c_26, c_r), "사업 소계 누계": ev(sr_idx, c_25, c_26, c_r), "사업 소계 월계": ev(sr_idx, m_25, m_26, m_r if m_r else c_r)}, {"누계": bc(c_25, c_26, c_r), "월계": bc(m_25, m_26, m_r)}

    bi_tot_k, bi_tot_c = parse_bi("종합")
    bi_sh_k, bi_sh_c = parse_bi("상해")
    bi_gw_k, bi_gw_c = parse_bi("광주")

    return {
        "sheets": total_sheets_count, "summary_chart": summary_chart, "calc_summary": calc_summary, "col_25": col_25, "col_26": col_26,
        "target_categories": target_cats, "part_data_cache": part_cache, "vendor_data_cache": vendor_cache,
        "bi_tot_k": bi_tot_k, "bi_tot_c": bi_tot_c, "bi_sh_k": bi_sh_k, "bi_sh_c": bi_sh_c, "bi_gw_k": bi_gw_k, "bi_gw_c": bi_gw_c
    }

parsed_data = load_and_parse_all_data(raw_bytes)
total_sheets_count = parsed_data["sheets"]
summary_chart, calc_summary = parsed_data["summary_chart"], parsed_data["calc_summary"]
col_25, col_26, target_categories = parsed_data["col_25"], parsed_data["col_26"], parsed_data["target_categories"]
part_data_cache, vendor_data_cache = parsed_data["part_data_cache"], parsed_data["vendor_data_cache"]

# =========================================================
# 11. KPI 영역 채우기 (URL 연동 상태 유지)
# =========================================================
with kpi_container:
    card_unit, display_period_name = t["unit"], t["periods"].get("전체 총계 누계", "전체 총계 누계")
    if page_menu.startswith("[BI_"):
        st.markdown("---")
        st.markdown(f"#### {t['period_select']}")
        bi_periods_keys = ["전체 총계 누계", "사업 소계 누계", "사업 소계 월계"]
        curr_p = query_params.get("period", None) if query_params.get("period") in bi_periods_keys else st.session_state.get("bi_period_mode", bi_periods_keys[0])
        for bp_key in bi_periods_keys:
            is_p_active = (curr_p == bp_key)
            disp_bp_key = t["periods"].get(bp_key, bp_key)
            st.markdown(f'<a href="?page={page_menu}&period={bp_key}&auth_ok=true&lang={current_lang_code}" class="{"sidebar-card-btn-active" if is_p_active else "sidebar-card-btn"}" style="font-weight: {"800" if is_p_active else "700"}; background-color: {"#003876" if is_p_active else "#F8FAFC"}; color: {"#FFFFFF" if is_p_active else "#0F172A"}; border: 1.5px solid {"#001E3D" if is_p_active else "#CBD5E1"};" target="_self">{disp_bp_key}</a>', unsafe_allow_html=True)
        if query_params.get("period") in bi_periods_keys: st.session_state["bi_period_mode"] = query_params.get("period")
        bi_period_mode = st.session_state.get("bi_period_mode", bi_periods_keys[0])
        display_period_name = t["periods"].get(bi_period_mode, bi_period_mode)

        t_kpi = parsed_data["bi_gw_k"] if "광주" in page_menu else (parsed_data["bi_sh_k"] if "상해" in page_menu else parsed_data["bi_tot_k"])
        bi_pack = t_kpi.get(bi_period_mode, t_kpi["전체 총계 누계"])
        total_25, total_26, diff_val, diff_rate = float(bi_pack["25"]), float(bi_pack["26"]), float(bi_pack["diff"]), float(bi_pack["rate"])
        
        bi_prefix = t.get("bi_gw", "BI_광주") if "광주" in page_menu else (t.get("bi_sh", "BI_상해") if "상해" in page_menu else t.get("bi_tot", "BI_종합"))
        card_sub_desc = f"{bi_prefix} [{display_period_name}]"
    else:
        all_biz_str = t.get("all_biz", "전체 사업 보기")
        sel_view = st.session_state.get("selected_biz_view", all_biz_str) if page_menu == "[접수기준] 사업별 실적 현황" else target_categories[0]
        
        logic_view = sel_view if sel_view in target_categories else "전체 사업 보기"
        
        if logic_view != "전체 사업 보기" and logic_view in target_categories:
            t_row = summary_chart[summary_chart["표준사업구분"] == logic_view]
            total_25, total_26 = float(t_row[col_25].sum()) if not t_row.empty else 0.0, float(t_row[col_26].sum()) if not t_row.empty else 0.0
            card_sub_desc = f"[{t['categories_map'].get(logic_view, logic_view)}] Total"
        else:
            total_25, total_26 = float(summary_chart[col_25].sum()), float(summary_chart[col_26].sum())
            card_sub_desc = "TOTAL Summary"
        diff_val = total_26 - total_25
        diff_rate = (diff_val / total_25 * 100) if total_25 != 0 else 0.0

# =========================================================
# 12. 상단 종합 KPI 카드 렌더링
# =========================================================
is_positive = diff_val >= 0
diff_color, badge_bg, diff_sign = ("#E11D48", "#FFE4E6", "+") if is_positive else ("#2563EB", "#DBEAFE", "")
c1, c2, c3, c4 = st.columns(4)

with c1: st.markdown(f'<div class="kpi-card" style="border-top-color: #64748B;"><div class="kpi-title">{t["kpi_25"]}</div><div class="kpi-num">{total_25:,.0f} <span style="font-size:14px; font-weight:normal; color:#64748B;">{card_unit}</span></div><div class="kpi-sub">{card_sub_desc}</div></div>', unsafe_allow_html=True)
with c2: st.markdown(f'<div class="kpi-card" style="border-top-color: #003876;"><div class="kpi-title">{t["kpi_26"]}</div><div class="kpi-num" style="color: #003876;">{total_26:,.0f} <span style="font-size:14px; font-weight:normal; color:#64748B;">{card_unit}</span></div><div class="kpi-sub">{card_sub_desc}</div></div>', unsafe_allow_html=True)
with c3: st.markdown(f'<div class="kpi-card" style="border-top-color: {diff_color};"><div class="kpi-title">{t["kpi_diff"]}</div><div class="kpi-num" style="color: {diff_color};">{diff_sign}{diff_val:,.0f} <span style="font-size:14px; font-weight:normal; color:#64748B;">{card_unit}</span></div><span class="kpi-badge" style="background-color: {badge_bg}; color: {diff_color};">{t["kpi_diff_sub"]}</span></div>', unsafe_allow_html=True)
with c4: st.markdown(f'<div class="kpi-card" style="border-top-color: {diff_color};"><div class="kpi-title">{t["kpi_rate"]}</div><div class="kpi-num" style="color: {diff_color};">{diff_sign}{diff_rate:0.1f}%</div><span class="kpi-badge" style="background-color: {badge_bg}; color: {diff_color};">{t["kpi_rate_sub"]}</span></div>', unsafe_allow_html=True)

st.write("")
st.markdown("---")

# =========================================================
# 13. 본문 렌더러 및 라우팅
# =========================================================
def wrap_text_for_axis(text, max_len=9):
    text_str = str(text)
    if len(text_str) <= max_len: return text_str
    if '(' in text_str and ')' in text_str:
        p = text_str.split('(')
        return p[0].strip() + "<br>(" + p[1].strip()
    words = text_str.split(' ')
    if len(words) > 1:
        mid = len(words) // 2
        return " ".join(words[:mid]) + "<br>" + " ".join(words[mid:])
    return text_str[:max_len] + "<br>" + text_str[max_len:]

def render_fullwidth_vertical_dashboard(title_top, title_bottom, table_title, data_df, x_col_name, cat_order):
    df = data_df.copy()
    c25_t = col_25 if col_25 in df.columns else "2025년 실적"
    c26_t = col_26 if col_26 in df.columns else "2026년 실적"
    if "2025년 실적" not in df.columns and c25_t in df.columns: df["2025년 실적"] = df[c25_t]
    if "2026년 실적" not in df.columns and c26_t in df.columns: df["2026년 실적"] = df[c26_t]
    if "증감액" not in df.columns: df["증감액"] = df["2026년 실적"] - df["2025년 실적"]
    if "증감률" not in df.columns or df["증감률"].isnull().all(): df["증감률"] = ((df["증감액"] / df["2025년 실적"].replace(0, pd.NA)) * 100).fillna(0.0)

    disp_x = f"{x_col_name}_wrapped"
    df[disp_x] = df[x_col_name].apply(lambda x: wrap_text_for_axis(x, 9))
    w_order = [wrap_text_for_axis(c, 9) for c in cat_order]

    def fmt_scale(val):
        av = abs(val)
        sign = "-" if val < 0 else ""
        if av >= 1e8: return f"{sign}{av/1e8:.1f}억"
        elif av >= 1e4: return f"{sign}{av/1e4:.0f}만"
        return f"{val:,.0f}"

    l25, l26, d_txt, d_col = [], [], [], []
    for _, r in df.iterrows():
        v25, v26, dv, rt = r["2025년 실적"], r["2026년 실적"], r["증감액"], r["증감률"]
        l25.append(f"<span style='font-size:13px; font-weight:700;'>{fmt_scale(v25)}</span>")
        rc = "#E11D48" if rt >= 0 else "#1D4ED8"
        l26.append(f"<span style='font-size:14px; font-weight:800;'>{fmt_scale(v26)}</span><br><span style='font-size:12px; font-weight:700; color:{rc};'>({'+' if rt>0 else ''}{rt:0.1f}%)</span>")
        dc = "#E11D48" if dv >= 0 else "#1D4ED8"
        d_txt.append(f"<span style='font-size:14px; font-weight:800; color:{dc};'>{'+' if dv>0 else ''}{fmt_scale(dv)}</span><br><span style='font-size:12px; font-weight:700; color:{dc};'>({'+' if rt>0 else ''}{rt:0.1f}%)</span>")
        d_col.append(dc)

    st.subheader(title_top)
    fig_b = go.Figure()
    fig_b.add_trace(go.Bar(x=df[disp_x], y=df["2025년 실적"], name="2025", marker=dict(color="#94A3B8", cornerradius=6), text=l25, textposition="outside"))
    fig_b.add_trace(go.Bar(x=df[disp_x], y=df["2026년 실적"], name="2026", marker=dict(color="#1D4ED8", cornerradius=6), text=l26, textposition="outside"))
    fig_b.update_layout(height=520, yaxis=dict(rangemode='tozero', gridcolor="#F1F5F9"), xaxis=dict(categoryorder='array', categoryarray=w_order), template="plotly_white", margin=dict(t=40, b=40, l=10, r=10))
    st.plotly_chart(fig_b, use_container_width=True)

    st.write("")
    st.markdown(f"##### {title_bottom}")
    fig_d = go.Figure()
    fig_d.add_trace(go.Bar(x=df[disp_x], y=df["증감액"], marker=dict(color=d_col, cornerradius=6), text=d_txt, textposition="outside"))
    fig_d.update_layout(height=420, yaxis=dict(gridcolor="#F1F5F9", zerolinecolor="#CBD5E1"), xaxis=dict(categoryorder='array', categoryarray=w_order), template="plotly_white", margin=dict(t=20, b=40, l=10, r=10))
    st.plotly_chart(fig_d, use_container_width=True)

    st.write("")
    st.markdown(f"##### 📋 {table_title}")
    st.dataframe(pd.DataFrame({x_col_name: df[x_col_name], "2025년": df["2025년 실적"], "2026년": df["2026년 실적"], "증감액": df["증감액"], "증감률(%)": df["증감률"]}), hide_index=True, use_container_width=True)

cur_page_disp = t["pages"].get(page_menu, page_menu)

if page_menu == "[접수기준] 종합 실적 현황":
    st.subheader(f"🥧 {cur_page_disp} - Share")
    biz_colors = {"글로벌 바이어": "#2563EB", "패션잡화": "#F59E0B", "GB": "#10B981", "제품평가": "#8B5CF6"}
    p_c1, p_c2 = st.columns(2)
    with p_c1:
        f_p25 = px.pie(summary_chart, names="표준사업구분", values=col_25, hole=0.6, title=t["pie_title_25"], category_orders={"표준사업구분": target_categories}, color="표준사업구분", color_discrete_map=biz_colors)
        f_p25.update_traces(textposition='inside', textinfo='label+percent', textfont=dict(size=14, color="#FFFFFF", weight="bold"))
        f_p25.update_layout(height=460, margin=dict(t=60, b=20, l=10, r=10), legend=dict(orientation="h", yanchor="bottom", y=-0.18, xanchor="center", x=0.5))
        st.plotly_chart(f_p25, use_container_width=True)
    with p_c2:
        f_p26 = px.pie(summary_chart, names="표준사업구분", values=col_26, hole=0.6, title=t["pie_title_26"], category_orders={"표준사업구분": target_categories}, color="표준사업구분", color_discrete_map=biz_colors)
        f_p26.update_traces(textposition='inside', textinfo='label+percent', textfont=dict(size=14, color="#FFFFFF", weight="bold"))
        f_p26.update_layout(height=460, margin=dict(t=60, b=20, l=10, r=10), legend=dict(orientation="h", yanchor="bottom", y=-0.18, xanchor="center", x=0.5))
        st.plotly_chart(f_p26, use_container_width=True)
    st.write("")
    st.markdown("---")
    render_fullwidth_vertical_dashboard(f"📌 {cur_page_disp}", "📈 Business Performance Diff & Growth Rate", "Summary Table", summary_chart, "표준사업구분", target_categories)

elif page_menu == "[접수기준] 사업별 실적 현황":
    all_biz_str = t.get("all_biz", "전체 사업 보기")
    sel_v = st.selectbox("Select Business:", [all_biz_str] + target_categories, key="sel_biz_v")
    df_v = summary_chart.copy() if sel_v == all_biz_str else calc_summary[calc_summary["표준사업구분"] == sel_v].copy()
    render_fullwidth_vertical_dashboard(f"🏢 {cur_page_disp} ({sel_v})", "📈 Detailed Performance Diff", "Detailed Summary Table", df_v.rename(columns={col_25: "2025년 실적", col_26: "2026년 실적"}), "표준사업구분" if sel_v == all_biz_str else "세부항목", target_categories if sel_v == all_biz_str else df_v["세부항목"].unique().tolist())

elif page_menu == "[접수기준] 바이어 실적 현황":
    sel_b = st.selectbox("Select Business:", target_categories, key="sel_t3_b")
    b_df = part_data_cache.get(sel_b, pd.DataFrame())
    if not b_df.empty:
        valid_b = b_df[~b_df["바이어명"].isin(["-", "", "NAN"])].sort_values(by="2026년 실적", ascending=False).reset_index(drop=True)
        render_fullwidth_vertical_dashboard(f"🤝 {cur_page_disp} ({sel_b})", "📈 Buyer Performance Diff", "Buyer Summary Table", valid_b.head(7), "바이어명", valid_b["바이어명"].head(7).tolist())

elif page_menu == "[접수기준] 협력사 실적 현황":
    c_b, c_v = st.columns([4, 6])
    with c_b: s_b = st.selectbox("Select Business:", target_categories, key="t4_b")
    v_df = vendor_data_cache.get(s_b, pd.DataFrame())
    all_vendor_str = t.get("all_vendor", "전체 협력사 보기")
    if not v_df.empty:
        v_sum = v_df.groupby("협력사명", as_index=False)[["2025년 실적", "2026년 실적"]].sum().sort_values(by="2026년 실적", ascending=False).reset_index(drop=True)
        with c_v: s_v = st.selectbox("Select Vendor:", [all_vendor_str] + v_sum["협력사명"].tolist(), key="t4_v")
        disp_v = v_sum.head(6) if s_v == all_vendor_str else v_df[v_df["협력사명"] == s_v].groupby("바이어명", as_index=False)[["2025년 실적", "2026년 실적"]].sum()
        perf_status_str = t.get("perf_status", " 실적 현황")
        render_fullwidth_vertical_dashboard(f"🏢 {s_v}{perf_status_str}", "📈 Vendor Performance Diff", "Vendor Summary Table", disp_v, "협력사명" if s_v == all_vendor_str else "바이어명", disp_v["협력사명" if s_v == all_vendor_str else "바이어명"].tolist())

elif page_menu.startswith("[BI_"):
    if "광주" in page_menu:
        chart_d = parsed_data["bi_gw_c"]["누계" if bi_period_mode == "전체 총계 누계" else ("월계" if "월계" in bi_period_mode else "누계")]
        center_title_prefix = f"🏭 [{t.get('bi_gw', 'BI_광주')}]"
    elif "상해" in page_menu:
        chart_d = parsed_data["bi_sh_c"]["누계" if bi_period_mode == "전체 총계 누계" else ("월계" if "월계" in bi_period_mode else "누계")]
        center_title_prefix = f"🏭 [{t.get('bi_sh', 'BI_상해')}]"
    else:
        chart_d = parsed_data["bi_tot_c"]["누계" if bi_period_mode == "전체 총계 누계" else ("월계" if "월계" in bi_period_mode else "누계")]
        center_title_prefix = f"📊 [{t.get('bi_tot', 'BI_종합')}]"

    mag_df = chart_d[chart_d["표준사업구분"].isin(MAGOK_CATEGORIES)].copy()
    och_df = chart_d[chart_d["표준사업구분"].isin(OCHANG_CATEGORIES)].copy()

    st.subheader(f"📍 {center_title_prefix} {t.get('center_compare_title', '마곡 본원 vs 오창 분원 거점별 실적 비교')} ({display_period_name})")
    
    mag_25 = mag_df["2025년 실적"].sum()
    mag_26 = mag_df["2026년 실적"].sum()
    och_25 = och_df["2025년 실적"].sum()
    och_26 = och_df["2026년 실적"].sum()

    c_pie_df = pd.DataFrame([
        {"거점구분": "마곡 본원 (Magok)", "2025년 실적": mag_25, "2026년 실적": mag_26},
        {"거점구분": "오창 분원 (Ochang)", "2025년 실적": och_25, "2026년 실적": och_26}
    ])
    
    cp1, cp2 = st.columns(2)
    ccol = {"마곡 본원 (Magok)": "#1D4ED8", "오창 분원 (Ochang)": "#10B981"}
    with cp1:
        fp1 = px.pie(c_pie_df, names="거점구분", values="2025년 실적", hole=0.6, title=t.get("center_pie_25", "2025년 거점별 실적 비중"), color="거점구분", color_discrete_map=ccol)
        fp1.update_traces(textposition='inside', textinfo='label+percent', textfont=dict(size=14, color="#FFFFFF", weight="bold"))
        fp1.update_layout(height=460, margin=dict(t=60, b=30, l=10, r=10), legend=dict(orientation="h", yanchor="bottom", y=-0.22, xanchor="center", x=0.5))
        st.plotly_chart(fp1, use_container_width=True)
    with cp2:
        fp2 = px.pie(c_pie_df, names="거점구분", values="2026년 실적", hole=0.6, title=t.get("center_pie_26", "2026년 거점별 실적 비중"), color="거점구분", color_discrete_map=ccol)
        fp2.update_traces(textposition='inside', textinfo='label+percent', textfont=dict(size=14, color="#FFFFFF", weight="bold"))
        fp2.update_layout(height=460, margin=dict(t=60, b=30, l=10, r=10), legend=dict(orientation="h", yanchor="bottom", y=-0.22, xanchor="center", x=0.5))
        st.plotly_chart(fp2, use_container_width=True)

    st.write("")
    st.markdown("---")
    
    center_comp_df = c_pie_df.copy()
    center_comp_df["증감액"] = center_comp_df["2026년 실적"] - center_comp_df["2025년 실적"]
    center_comp_df["증감률"] = ((center_comp_df["증감액"] / center_comp_df["2025년 실적"].replace(0, pd.NA)) * 100).fillna(0.0)
    render_fullwidth_vertical_dashboard(f"{center_title_prefix} {t.get('center_growth', '마곡 본원 vs 오창 분원 요약 비교')}", "📈 Center Growth Comparison", "Magok & Ochang Summary Table", center_comp_df, "거점구분", ["마곡 본원 (Magok)", "오창 분원 (Ochang)"])

    st.write("")
    st.markdown("---")
    render_fullwidth_vertical_dashboard(f"🏛️ {t.get('magok_title', '마곡 본원 세부 사업별 실적 현황')}", "📈 Magok Diff", "Magok Table", mag_df, "표준사업구분", MAGOK_CATEGORIES)
    
    st.write("")
    st.markdown("---")
    render_fullwidth_vertical_dashboard(f"🏭 {t.get('ochang_title', '오창 분원 세부 사업별 실적 현황')}", "📈 Ochang Diff", "Ochang Table", och_df, "표준사업구분", OCHANG_CATEGORIES)

    st.write("")
    st.markdown("---")
    render_fullwidth_vertical_dashboard(f"{center_title_prefix} {t.get('all_cat_title', '전체 12대 사업별 상세 실적 현황')}", "📈 All Categories Performance Diff", "All Categories Detailed Summary Table", chart_d, "표준사업구분", FULL_BI_CATEGORIES)
