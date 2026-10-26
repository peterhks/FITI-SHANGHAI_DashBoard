import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import io
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
# 2. 사용자 권한 DB (Supabase 클라우드 영구 저장소 완벽 연동)
# =========================================================
try:
    from supabase import create_client
    SUPABASE_URL = st.secrets["SUPABASE_URL"]
    SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
except Exception:
    supabase = None

def load_user_db():
    default_db = {"kshan@fiti.re.kr": {"pw": "fb09010552", "role": "admin", "name": "관리자"}}
    if supabase:
        try:
            res = supabase.table("fiti_users").select("*").execute()
            if res.data:
                db = {}
                for row in res.data:
                    db[row['email']] = {"pw": row['pw'], "role": row['role'], "name": row['name']}
                return db
        except Exception:
            pass
    return default_db

def save_user_db(db_data):
    if supabase:
        try:
            for email, info in db_data.items():
                supabase.table("fiti_users").upsert({
                    "email": email,
                    "pw": info["pw"],
                    "role": info["role"],
                    "name": info["name"]
                }).execute()
        except Exception:
            pass

def delete_user_db(email):
    if supabase:
        try:
            supabase.table("fiti_users").delete().eq("email", email).execute()
        except Exception:
            pass

def load_login_history():
    if supabase:
        try:
            res = supabase.table("login_history").select("time, email, name, status").execute()
            if res.data:
                return res.data
        except Exception:
            pass
    return []

def append_login_history(record):
    if supabase:
        try:
            supabase.table("login_history").insert([record]).execute()
        except Exception:
            pass

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
        "login_log": "로그인 기록 로그",
        "date_range": "날짜 및 기간 지정",
        "download_csv": "📥 다운로드 (CSV)",
        "no_record": "기록 없음",
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
        "login_log": "Login History Log",
        "date_range": "Select Date Range",
        "download_csv": "📥 Download (CSV)",
        "no_record": "No records",
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
# 4. 스타일 및 디자인 (텍스트 노출 방지 처리 완벽 적용)
# =========================================================
def inject_css(css_string):
    try:
        st.html(f"<style>{css_string}</style>")
    except Exception:
        st.markdown(f"<style>{css_string}</style>", unsafe_allow_html=True)

main_css = """
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable.min.css');
html, body, [class*="css"] { font-family: 'Pretendard Variable', Pretendard, -apple-system, BlinkMacSystemFont, 'Apple SD Gothic Neo', system-ui, Roboto, sans-serif !important; }
section[data-testid="stSidebar"] { padding-top: 0rem !important; }
section[data-testid="stSidebar"] div.block-container { padding-top: 0.1rem !important; padding-bottom: 0.4rem !important; }
section[data-testid="stSidebar"] div.stExpander { margin-bottom: 0.1rem !important; }
section[data-testid="stSidebar"] hr { margin: 0.2rem 0 !important; }
div[data-testid="stSidebar"] div.stExpander button
