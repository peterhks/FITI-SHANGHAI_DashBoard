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
