# =========================================================
# 2. FITI 상해지사 공식 상단 배너 헤더
# =========================================================
st.markdown("""
<div class="fiti-header">
    <div class="fiti-logo-area">
        <div class="fiti-logo-text">FITI</div>
        <div class="fiti-sub-logo">
            <span class="fiti-sub-logo-kr">FITI 상해시험연구원</span>
            <span class="fiti-sub-logo-en">FITI Shanghai Branch</span>
        </div>
    </div>
    <div class="fiti-title-area">
        <div class="fiti-main-title">FITI 상해지사 실적 종합 분석 시스템</div>
        <div class="fiti-en-title">Shanghai Branch Business Performance & Testing Analytics System</div>
        <div class="fiti-dept-title">사업팀 (Business Operations Team) · 제품평가팀 (Inspection Team)</div>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================
# 3. 사이트 목적 맞춤형 사용 안내 (Expander)
# =========================================================
with st.expander("ℹ️ 사용 안내", expanded=True):
    st.markdown("""
    * **실적 데이터 연동**: 저장소 내 기본 실적 파일(`performance_260825.xlsx`)이 자동 반영되며, 좌측 패널을 통해 최신 접수 실적(Excel/CSV)을 직접 업로드할 수 있습니다.
    * **사업 영역별 다차원 집계**: 중국 국가표준(GB), 한국 수출 KC 인증, 글로벌 바이어 매뉴얼 시험 및 완제품 공장 검사(제품평가) 실적을 고객사·업무구분별로 통합 분석합니다.
    * **시계열 추이 및 수요 예측**: 일자별 수수료 매출 및 성적서 발급 추이를 확인하고, 통계 모델(ARIMA $p, d, q$)을 통해 향후 시험 접수 수요를 선제적으로 예측합니다.
    * **고객사 기여도 및 이상 징후 진단**: 주요 고객사(브랜드)별 기여도 순위(Funnel)와 통계적 관리한계(±2σ)를 이탈한 이상 실적 패턴을 자동으로 추출하여 리포트합니다.
    """)
