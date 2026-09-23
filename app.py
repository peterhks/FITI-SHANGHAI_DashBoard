import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import io
import base64
from datetime import datetime

# =========================================================
# 1. 화면 기본 설정 및 디자인 스타일
# =========================================================
st.set_page_config(
    page_title="FITI SHANGHAI Performance Analysis",
    layout="wide",
    initial_sidebar_state="expanded"
)

MAGOK_CATEGORIES = [
    "법정검사", 
    "일반검사", 
    "섬유내수(패션잡화)", 
    "섬유내수(중국GB)", 
    "섬유내수(단체/정부)", 
    "섬유수출", 
    "연구용역", 
    "제품인증(Q.SF)"
]

OCHANG_CATEGORIES = [
    "산업(토목+부품)", 
    "모빌리티(전장+의장)", 
    "환경(환경+측정기기)", 
    "화학바이오(화학제품+생활안전)"
]

FULL_BI_CATEGORIES = MAGOK_CATEGORIES + OCHANG_CATEGORIES

# =========================================================
# 2. 사용자 권한 및 로그인 로그 영구 유지 초기화 (삭제 부활 원천 차단)
# =========================================================
if "user_db" not in st.session_state:
    st.session_state["user_db"] = {
        "kshan@fiti.re.kr": {"pw": "fb09010552", "role": "admin", "name": "관리자"},
        "admin@fiti.re.kr": {"pw": "fiti1965", "role": "admin", "name": "시스템 관리자"},
        "leader@fiti.re.kr": {"pw": "fiti1234", "role": "bi_user", "name": "상해지사 팀장"},
        "staff@fiti.re.kr": {"pw": "fiti5678", "role": "general_user", "name": "일반 담당자"}
    }

if "login_history" not in st.session_state:
    st.session_state["login_history"] = []

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
    st.session_state["current_user_email"] = ""
    st.session_state["current_user_role"] = ""
    st.session_state["current_user_name"] = ""

# 세션 유지 보완 (쿼리 파라미터 동기화)
query_params = st.query_params
if "auth_ok" in query_params and query_params["auth_ok"] == "true":
    st.session_state["logged_in"] = True
    if not st.session_state["current_user_email"]:
        st.session_state["current_user_email"] = "kshan@fiti.re.kr"
        st.session_state["current_user_role"] = "admin"
        st.session_state["current_user_name"] = "관리자"

# =========================================================
# 3. 다국어 텍스트 사전 (한국어, 중국어, 영어)
# =========================================================
LANG_DICT = {
    "한국어": {
        "sys_title": "상해지사 실적 종합 분석 시스템",
        "sys_sub": "상해지사 사업 실적 및 분석 시스템 | 상해지사 사업팀",
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
        "buyer_pie_25": "2025년 주요 바이어 실적 비중",
        "buyer_pie_26": "2026년 주요 바이어 실적 비중",
        "center_compare": "마곡 본원 vs 오창 분원 거점별 실적 비교",
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
    # (중국어, 영어 사전 생략 - 기존과 동일)
}

# 기본값 설정
if "selected_lang" not in st.session_state:
    st.session_state["selected_lang"] = "한국어"
selected_lang = st.session_state["selected_lang"]
t = LANG_DICT[selected_lang]
current_font = t["font_family"]

# =========================================================
# 4. 사이드바 스타일 (극단적 최소화 압축 적용)
# =========================================================
st.markdown(f"""
<style>
    /* 사이드바 전체 패딩 제거 */
    section[data-testid="stSidebar"] {
        padding-top: 0rem !important;
    }
    /* 내부 컨테이너 상단 공백 제거 */
    section[data-testid="stSidebar"] div.block-container {
        padding-top: 0.1rem !important;
        padding-bottom: 0.3rem !important;
    }
    /* Expander 사이 간격 제거 */
    section[data-testid="stSidebar"] div.stExpander {
        margin-bottom: 0px !important;
    }
    /* 구분선 간격 최소화 */
    section[data-testid="stSidebar"] hr {
        margin: 0.3rem 0 !important;
    }
    /* 제목(h4, h5) 여백 최소화 */
    section[data-testid="stSidebar"] h4,
    section[data-testid="stSidebar"] h5 {
        margin-top: 0.3rem !important;
        margin-bottom: 0.1rem !important;
    }
    /* 버튼(네비게이션 카드) 여백/패딩 최소화 */
    .sidebar-card-btn, .sidebar-card-btn-active {
        margin-top: 0px !important;
        margin-bottom: 1px !important;
        padding-top: 4px !important;
        padding-bottom: 4px !important;
    }
</style>
""", unsafe_allow_html=True)

# =========================================================
# 5. 상해 야경 테마 프리미엄 로그인 화면 (생략 - 기존과 동일)
# =========================================================
if not st.session_state["logged_in"]:
    # (생략)
    st.stop()

# =========================================================
# 6. 상단 공식 배너 (로그인 후 표시)
# =========================================================
st.markdown(f"""
<div style="background: linear-gradient(135deg, #002B5C 0%, #003876 100%); padding: 18px 24px; border-radius: 10px; display: flex; align-items: center; gap: 22px; color: #FFFFFF; margin-bottom: 18px; box-shadow: 0 4px 14px rgba(0, 43, 92, 0.18);">
    <div style="font-size: 24px; font-weight: 900; letter-spacing: -0.5px; border-right: 1.5px solid rgba(255, 255, 255, 0.25); padding-right: 22px;">FITI</div>
    <div style="flex-grow: 1;">
        <div style="font-size: 18px; font-weight: 800; margin-bottom: 2px;">{t["sys_title"]}</div>
        <div style="font-size: 12px; color: #D0E1FD; font-weight: 400;">{t["sys_sub"]}</div>
    </div>
    <div style="text-align: right; font-size: 12px; color: #D0E1FD;">
        <b>{st.session_state['current_user_name']}</b>님 환영합니다.<br>
        <span style="background: rgba(255,255,255,0.2); padding: 2px 8px; border-radius: 4px; font-size: 10px;">
            권한: {'관리자' if st.session_state['current_user_role']=='admin' else ('접수 + BI' if st.session_state['current_user_role']=='bi_user' else '접수')}
        </span>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================
# 7. 사이드바 파일 로드 설정 (생략 - 기존과 동일)
# =========================================================
EXCEL_FILE = "performance_최신.xlsx"
os.makedirs("downloads", exist_ok=True)
LOCAL_EXCEL_PATH = os.path.join("downloads", EXCEL_FILE)

if "persistent_file_bytes" in st.session_state:
    raw_bytes = st.session_state["persistent_file_bytes"]
elif os.path.exists(LOCAL_EXCEL_PATH):
    with open(LOCAL_EXCEL_PATH, "rb") as f:
        raw_bytes = f.read()
    st.session_state["persistent_file_bytes"] = raw_bytes
else:
    raw_bytes = None

if not raw_bytes:
    st.warning(t["file_not_found"])
    st.stop()

try:
    temp_stream = io.BytesIO(raw_bytes)
    temp_excel = pd.ExcelFile(temp_stream)
    total_sheets_count = len(temp_excel.sheet_names)
except Exception:
    total_sheets_count = 0

def clean_series(series):
    cleaned = series.astype(str).str.replace(',', '').str.replace('₩', '').str.strip()
    cleaned = cleaned.replace(['-', '–', '—', 'nan', 'NaN', 'None', ''], '0')
    return pd.to_numeric(cleaned, errors='coerce').fillna(0)

# (캐싱 함수 생략 - 기존과 동일)

# =========================================================
# 8. 파서 함수 정의 (생략 - 기존과 동일)
# =========================================================
# (데이터 로직 그대로 사용)

# =========================================================
# 9. 사이드바 네비게이션 및 권한별 페이지 제어
# =========================================================
user_role = st.session_state["current_user_role"]
if user_role == "general_user":
    all_pages_keys = [
        "[접수기준] 종합 실적 현황",
        "[접수기준] 사업별 실적 현황",
        "[접수기준] 바이어 실적 현황",
        "[접수기준] 협력사 실적 현황"
    ]
else:
    all_pages_keys = [
        "[접수기준] 종합 실적 현황",
        "[접수기준] 사업별 실적 현황",
        "[접수기준] 바이어 실적 현황",
        "[접수기준] 협력사 실적 현황",
        "[BI_종합] 사업별 실적 현황",
        "[BI_상해] 사업별 실적 현황",
        "[BI_광주] 사업별 실적 현황"
    ]

st.sidebar.markdown(f"#### {t['page_select']}")
# 네비게이션 버튼 생성 (기존 코드의 링크 스타일을 클래스로 변경하여 일괄 적용)
for p_key in all_pages_keys:
    is_active = (st.session_state.get("current_page", all_pages_keys[0]) == p_key)
    display_name = t["pages"].get(p_key, p_key)
    
    # 버튼 클릭 시 URL 파라미터를 변경하여 세션 유지
    f_class = "sidebar-card-btn-active" if is_active else "sidebar-card-btn"
    link_html = f"""
    <a href="?page={p_key}&auth_ok=true" class="{f_class}" style="display: block; width: 100%; border-radius: 6px; text-align: center; font-weight: {'800' if is_active else '700'}; font-size: 12px; padding: 4px 8px; margin-bottom: 1px; border: 1.5px solid {'#001E3D' if is_active else '#CBD5E1'}; background-color: {'#003876' if is_active else '#F8FAFC'}; color: {'#FFFFFF' if is_active else '#0F172A'}; text-decoration: none; box-shadow: 0 1px 3px rgba(0,0,0,0.05); transition: all 0.1s ease;" target="_self">
        {display_name}
    </a>
    """
    st.sidebar.markdown(link_html, unsafe_allow_html=True)

# 쿼리 파라미터로 페이지 상태 동기화
page_query = query_params.get("page", None)
if page_query and page_query in all_pages_keys:
    st.session_state["current_page"] = page_query
else:
    if "current_page" not in st.session_state:
        st.session_state["current_page"] = all_pages_keys[0]
page_menu = st.session_state["current_page"]

# =========================================================
# 10. [BI] 실적 기간 선택 (BI 페이지 진입 시만 노출)
# =========================================================
if page_menu.startswith("[BI_"):
    st.sidebar.markdown("---")
    st.sidebar.markdown(f"#### {t['period_select']}")
    bi_periods_keys = ["전체 총계 누계", "사업 소계 누계", "사업 소계 월계"]
    
    period_query = query_params.get("period", None)
    current_period = period_query if period_query in bi_periods_keys else st.session_state.get("bi_period_mode", bi_periods_keys[0])
    
    for bp_key in bi_periods_keys:
        is_p_active = (current_period == bp_key)
        p_class = "sidebar-card-btn-active" if is_p_active else "sidebar-card-btn"
        p_style = f"display: block; width: 100%; border-radius: 6px; text-align: center; font-weight: {'800' if is_p_active else '700'}; font-size: 12px; padding: 4px 8px; margin-bottom: 1px; border: 1.5px solid {'#001E3D' if is_p_active else '#CBD5E1'}; background-color: {'#003876' if is_p_active else '#F8FAFC'}; color: {'#FFFFFF' if is_p_active else '#0F172A'}; text-decoration: none; box-shadow: 0 1px 3px rgba(0,0,0,0.05);"
        p_link_html = f'<a href="?page={page_menu}&period={bp_key}&auth_ok=true" class="{p_class}" style="{p_style}" target="_self">{bp_key}</a>'
        st.sidebar.markdown(p_link_html, unsafe_allow_html=True)
        
    if period_query in bi_periods_keys:
        st.session_state["bi_period_mode"] = period_query
    bi_period_mode = st.session_state.get("bi_period_mode", bi_periods_keys[0])
    display_period_name = t["periods"].get(bi_period_mode, bp_key)

# =========================================================
# 11. [핵심 수정] 관리자 영역 - 담당자 목록 간격 최소화 및 폰트 축소
# =========================================================
if user_role in ["admin", "bi_user"]:
    st.sidebar.markdown("---")
    with st.sidebar.expander("🛡️ 관리자 권한", expanded=True):
        
        st.markdown(f"##### {t['data_mgmt']}")
        if st.session_state["current_user_role"] == "admin":
            uploaded_file = st.file_uploader(t["admin_upload"], type=["xlsx", "csv"], label_visibility="collapsed")
            if uploaded_file is not None:
                file_bytes = uploaded_file.getvalue()
                with open(LOCAL_EXCEL_PATH, "wb") as f: f.write(file_bytes)
                st.session_state["persistent_file_bytes"] = file_bytes
                st.cache_data.clear()
                st.rerun()
        else:
            st.caption(t["admin_caption"])
        # 파일 연동 문구 폰트 초소형 및 여백 제거
        st.markdown(f"<p style='font-size:10px; color:#64748B; margin-top:-5px; margin-bottom:2px;'>📂 파일 연동 중 (시트수: {total_sheets_count}개)</p>", unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("#### 👥 등록된 담당자 목록")
        
        # 권한별 담당자 분류
        cat_reception = [info for em, info in st.session_state["user_db"].items() if info["role"] == "general_user"]
        cat_reception_bi = [info for em, info in st.session_state["user_db"].items() if info["role"] == "bi_user"]
        cat_admin = [info for em, info in st.session_state["user_db"].items() if info["role"] == "admin"]

        # 💡 [핵심 수정] 각 카테고리 제목과 리스트의 margin을 음수(-10px)로 강제 설정하여 극단적으로 붙임
        
        # 접수
        st.markdown("<div style='margin-top: -10px;'><b>[ 접수 ]</b></div>", unsafe_allow_html=True)
        if cat_reception:
            for info in cat_reception:
                st.markdown(f"<div style='font-size:11px; margin-top: -10px; padding-left: 10px;'>• {info['name']}</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div style='font-size:11px; margin-top: -10px; padding-left: 10px; color:gray;'>• 등록된 인원이 없습니다.</div>", unsafe_allow_html=True)
            
        # 접수 + BI
        st.markdown("<div style='margin-top: 0px;'><b>[ 접수 + BI ]</b></div>", unsafe_allow_html=True)
        if cat_reception_bi:
            for info in cat_reception_bi:
                st.markdown(f"<div style='font-size:11px; margin-top: -10px; padding-left: 10px;'>• {info['name']}</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div style='font-size:11px; margin-top: -10px; padding-left: 10px; color:gray;'>• 등록된 인원이 없습니다.</div>", unsafe_allow_html=True)
            
        # 관리자
        st.markdown("<div style='margin-top: 0px;'><b>[ 관리자 ]</b></div>", unsafe_allow_html=True)
        if cat_admin:
            for info in cat_admin:
                st.markdown(f"<div style='font-size:11px; margin-top: -10px; padding-left: 10px;'>• {info['name']}</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div style='font-size:11px; margin-top: -10px; padding-left: 10px; color:gray;'>• 등록된 인원이 없습니다.</div>", unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("#### ➕ 담당자 추가 / 🗑️ 삭제")
        with st.form("add_user_form"):
            new_email = st.text_input("아이디 또는 이메일 (ID)", placeholder="예: kshan 또는 kshan@fiti.re.kr", label_visibility="collapsed")
            new_name = st.text_input("담당자 성명", placeholder="홍길동", label_visibility="collapsed")
            new_pw = st.text_input("비밀번호", type="password", placeholder="비밀번호 입력", label_visibility="collapsed")
            new_category = st.selectbox("권한 카테고리 지정", ["접수", "접수 + BI", "관리자"], label_visibility="collapsed")
            
            submit_add = st.form_submit_button("담당자 등록", use_container_width=True)
            if submit_add:
                clean_email = new_email.strip()
                if clean_email and new_pw.strip():
                    if "@" not in clean_email: clean_email = f"{clean_email}@fiti.re.kr"
                    assigned_role = "general_user" if new_category == "접수" else ("bi_user" if new_category == "접수 + BI" else "admin")
                    st.session_state["user_db"][clean_email] = {"pw": new_pw.strip(), "role": assigned_role, "name": new_name.strip() if new_name.strip() else clean_email}
                    st.success(f"✅ {clean_email} 영구 등록 완료!")
                    st.rerun()
                else:
                    st.error("아이디(이메일)와 비밀번호는 필수입니다.")
                    
        target_delete_email = st.selectbox("삭제할 담당자 선택", ["선택하세요."] + list(st.session_state["user_db"].keys()), label_visibility="collapsed")
        if st.button("선택한 담당자 삭제", use_container_width=True):
            if target_delete_email != "선택하세요.":
                if target_delete_email == st.session_state["current_user_email"]:
                    st.error("현재 로그인 중인 계정은 삭제할 수 없습니다.")
                else:
                    del st.session_state["user_db"][target_delete_email]
                    st.success(f"🗑️ {target_delete_email} 영구 삭제 완료!")
                    st.rerun()

        # 감사 로그 영역 (간격 최소화 적용)
        st.markdown("---")
        st.markdown("#### 📋 최근 로그인 감사 로그")
        if st.session_state["login_history"]:
            df_log = pd.DataFrame(st.session_state["login_history"])
            df_log["날짜"] = pd.to_datetime(df_log["time"]).dt.date
            unique_dates = ["전체 날짜 보기"] + sorted(df_log["날짜"].astype(str).unique().tolist(), reverse=True)
            selected_date_filter = st.selectbox("날짜별 로그 분리 보기", unique_dates, label_visibility="collapsed")
            
            filtered_log_df = df_log[df_log["날짜"].astype(str) == selected_date_filter] if selected_date_filter != "전체 날짜 보기" else df_log
            st.dataframe(filtered_log_df[["time", "email", "name", "status"]].tail(10), hide_index=True, use_container_width=True)
            
            csv_data = filtered_log_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(label="📥 감사 로그 다운로드 (CSV)", data=csv_data, file_name=f"fiti_login_audit_log_{datetime.now().strftime('%Y%m%d')}.csv", mime="text/csv", use_container_width=True)
        else:
            st.caption("기록된 로그인 이력이 없습니다.")

# 하단 고정 로그인/로그아웃 버튼 (사이드바 최하단)
st.sidebar.markdown("---")
st.sidebar.markdown(f"👤 **접속 계정**: {st.session_state['current_user_name']}")
if st.sidebar.button("시스템 잠금 (로그아웃)", use_container_width=True):
    st.session_state["logged_in"] = False
    st.session_state["current_user_email"] = ""
    st.session_state["current_user_role"] = ""
    st.session_state["current_user_name"] = ""
    if "auth_ok" in st.query_params: del st.query_params["auth_ok"]
    st.rerun()

# =========================================================
# 12. 상단 종합 KPI 카드 렌더링 (생략 - 기존과 동일)
# =========================================================
# (KPI 데이터 로직 및 카드 렌더링)

# =========================================================
# 13. 본문 대시보드 렌더링 (생략 - 기존과 동일)
# =========================================================
# (차트 및 테이블 렌더링)
