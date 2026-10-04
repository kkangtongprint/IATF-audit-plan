import streamlit as st
import pandas as pd
import io
import math
import re
import os
from datetime import datetime, timedelta, date

try:
    import docx
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

st.set_page_config(page_title="IATF 16949 CARA 심사계획서 생성기", page_icon="📋", layout="wide")
APP_VERSION = "2026-10-04-v4"

# ──────────────────────────────────────────────
# 상수
# ──────────────────────────────────────────────
HOURS_PER_MD   = 8.0
MAX_DAILY_HOURS = 10.0
MIN_MFG_RATIO  = 0.31
MAX_MFG_RATIO  = 0.34

EAC_OPTIONS = ["EAC 12","EAC 14","EAC 17","EAC 19","EAC 22","EAC 29","EAC 34"]

AUDITOR_INFO_FILENAME = "IATF 심사원 정보.xlsx"

def _load_auditor_db():
    """앱 폴더의 심사원 정보 Excel을 기준으로 심사원 목록을 구성합니다."""
    candidates = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), AUDITOR_INFO_FILENAME),
        os.path.join(os.getcwd(), AUDITOR_INFO_FILENAME),
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                df = pd.read_excel(path, dtype=str).fillna("")
                cols = {str(c).strip(): c for c in df.columns}
                kr_col = cols.get("국문이름")
                en_col = cols.get("영문이름")
                if kr_col and en_col:
                    rows = []
                    for _, r in df.iterrows():
                        kr = str(r[kr_col]).strip()
                        en = str(r[en_col]).strip()
                        if kr:
                            rows.append({"kr": kr, "en": en})
                    if rows:
                        return rows
            except Exception:
                pass
    # Excel을 찾지 못한 경우의 최소 호환 목록
    return [
        {"kr":"강병수","en":"Byeongsu Kang 강병수"},
        {"kr":"이승찬","en":"Seung-Chan Lee"},
        {"kr":"이수열","en":"Soo-Yeoul Lee"},
        {"kr":"정유근","en":"You-Gun Jung"},
        {"kr":"정영진","en":"Young Jin Jeong"},
        {"kr":"강성하","en":"Seong-Ha Kang"},
        {"kr":"강기원","en":"Gi-Won Kang"},
        {"kr":"전경준","en":"Gyeong-Jun Jeon"},
    ]

AUDITOR_DB = _load_auditor_db()
KR_TO_EN = {a["kr"]: a["en"] for a in AUDITOR_DB}
KR_NAMES = ["--심사원 선택--"] + [a["kr"] for a in AUDITOR_DB]

# 이전 부적합 검증용 조항 번호: ISO 9001:2015 + IATF 16949:2016의 주요/세부 조항 번호를 선택할 수 있도록 구성
ISO_CLAUSES = [
    "ISO 9001 - 4.1", "ISO 9001 - 4.2", "ISO 9001 - 4.3", "ISO 9001 - 4.4", "ISO 9001 - 4.4.1", "ISO 9001 - 4.4.2",
    "ISO 9001 - 5.1", "ISO 9001 - 5.1.1", "ISO 9001 - 5.1.2", "ISO 9001 - 5.2", "ISO 9001 - 5.2.1", "ISO 9001 - 5.2.2", "ISO 9001 - 5.3",
    "ISO 9001 - 6.1", "ISO 9001 - 6.1.1", "ISO 9001 - 6.1.2", "ISO 9001 - 6.2", "ISO 9001 - 6.2.1", "ISO 9001 - 6.2.2", "ISO 9001 - 6.3",
    "ISO 9001 - 7.1", "ISO 9001 - 7.1.1", "ISO 9001 - 7.1.2", "ISO 9001 - 7.1.3", "ISO 9001 - 7.1.4", "ISO 9001 - 7.1.5", "ISO 9001 - 7.1.5.1", "ISO 9001 - 7.1.5.2", "ISO 9001 - 7.1.6", "ISO 9001 - 7.2", "ISO 9001 - 7.3", "ISO 9001 - 7.4", "ISO 9001 - 7.5", "ISO 9001 - 7.5.1", "ISO 9001 - 7.5.2", "ISO 9001 - 7.5.3",
    "ISO 9001 - 8.1", "ISO 9001 - 8.2", "ISO 9001 - 8.2.1", "ISO 9001 - 8.2.2", "ISO 9001 - 8.2.3", "ISO 9001 - 8.2.4", "ISO 9001 - 8.3", "ISO 9001 - 8.3.1", "ISO 9001 - 8.3.2", "ISO 9001 - 8.3.3", "ISO 9001 - 8.3.4", "ISO 9001 - 8.3.5", "ISO 9001 - 8.3.6", "ISO 9001 - 8.4", "ISO 9001 - 8.4.1", "ISO 9001 - 8.4.2", "ISO 9001 - 8.4.3", "ISO 9001 - 8.5", "ISO 9001 - 8.5.1", "ISO 9001 - 8.5.2", "ISO 9001 - 8.5.3", "ISO 9001 - 8.5.4", "ISO 9001 - 8.5.5", "ISO 9001 - 8.5.6", "ISO 9001 - 8.6", "ISO 9001 - 8.7", "ISO 9001 - 8.7.1", "ISO 9001 - 8.7.2",
    "ISO 9001 - 9.1", "ISO 9001 - 9.1.1", "ISO 9001 - 9.1.2", "ISO 9001 - 9.1.3", "ISO 9001 - 9.2", "ISO 9001 - 9.2.1", "ISO 9001 - 9.2.2", "ISO 9001 - 9.3", "ISO 9001 - 9.3.1", "ISO 9001 - 9.3.2", "ISO 9001 - 9.3.3",
    "ISO 9001 - 10.1", "ISO 9001 - 10.2", "ISO 9001 - 10.3",
]
IATF_CLAUSES = [
    "IATF 16949 - 4.1", "IATF 16949 - 4.2", "IATF 16949 - 4.3", "IATF 16949 - 4.3.1", "IATF 16949 - 4.3.2", "IATF 16949 - 4.4", "IATF 16949 - 4.4.1", "IATF 16949 - 4.4.1.1", "IATF 16949 - 4.4.1.2", "IATF 16949 - 4.4.1.3", "IATF 16949 - 4.4.1.4", "IATF 16949 - 4.4.2",
    "IATF 16949 - 5.1", "IATF 16949 - 5.1.1", "IATF 16949 - 5.1.1.1", "IATF 16949 - 5.1.1.2", "IATF 16949 - 5.1.1.3", "IATF 16949 - 5.1.2", "IATF 16949 - 5.1.2.1", "IATF 16949 - 5.2", "IATF 16949 - 5.2.1", "IATF 16949 - 5.2.2", "IATF 16949 - 5.3", "IATF 16949 - 5.3.1",
    "IATF 16949 - 6.1", "IATF 16949 - 6.1.1", "IATF 16949 - 6.1.2", "IATF 16949 - 6.1.2.1", "IATF 16949 - 6.1.2.2", "IATF 16949 - 6.1.2.3", "IATF 16949 - 6.1.2.4", "IATF 16949 - 6.2", "IATF 16949 - 6.2.1", "IATF 16949 - 6.2.2", "IATF 16949 - 6.2.2.1", "IATF 16949 - 6.2.2.2",
    "IATF 16949 - 7.1", "IATF 16949 - 7.1.1", "IATF 16949 - 7.1.2", "IATF 16949 - 7.1.3", "IATF 16949 - 7.1.3.1", "IATF 16949 - 7.1.3.2", "IATF 16949 - 7.1.4", "IATF 16949 - 7.1.4.1", "IATF 16949 - 7.1.4.2", "IATF 16949 - 7.1.5", "IATF 16949 - 7.1.5.1", "IATF 16949 - 7.1.5.1.1", "IATF 16949 - 7.1.5.2", "IATF 16949 - 7.1.5.2.1", "IATF 16949 - 7.1.5.3", "IATF 16949 - 7.1.5.3.1", "IATF 16949 - 7.1.5.3.2", "IATF 16949 - 7.1.6", "IATF 16949 - 7.2", "IATF 16949 - 7.2.1", "IATF 16949 - 7.2.2", "IATF 16949 - 7.2.3", "IATF 16949 - 7.3", "IATF 16949 - 7.3.1", "IATF 16949 - 7.3.2", "IATF 16949 - 7.3.2.1", "IATF 16949 - 7.3.3", "IATF 16949 - 7.3.3.1", "IATF 16949 - 7.4", "IATF 16949 - 7.5", "IATF 16949 - 7.5.1", "IATF 16949 - 7.5.1.1", "IATF 16949 - 7.5.1.2", "IATF 16949 - 7.5.1.3", "IATF 16949 - 7.5.1.4", "IATF 16949 - 7.5.1.5", "IATF 16949 - 7.5.1.6", "IATF 16949 - 7.5.1.7", "IATF 16949 - 7.5.1.8", "IATF 16949 - 7.5.1.9", "IATF 16949 - 7.5.1.10", "IATF 16949 - 7.5.2", "IATF 16949 - 7.5.3", "IATF 16949 - 7.5.3.1", "IATF 16949 - 7.5.3.2", "IATF 16949 - 7.5.3.2.1", "IATF 16949 - 7.5.3.2.2",
    "IATF 16949 - 8.1", "IATF 16949 - 8.1.1", "IATF 16949 - 8.1.2", "IATF 16949 - 8.1.3", "IATF 16949 - 8.1.4", "IATF 16949 - 8.2", "IATF 16949 - 8.2.1", "IATF 16949 - 8.2.1.1", "IATF 16949 - 8.2.2", "IATF 16949 - 8.2.2.1", "IATF 16949 - 8.2.3", "IATF 16949 - 8.2.3.1", "IATF 16949 - 8.2.3.1.1", "IATF 16949 - 8.2.4", "IATF 16949 - 8.3", "IATF 16949 - 8.3.1", "IATF 16949 - 8.3.1.1", "IATF 16949 - 8.3.2", "IATF 16949 - 8.3.2.1", "IATF 16949 - 8.3.3", "IATF 16949 - 8.3.3.1", "IATF 16949 - 8.3.3.2", "IATF 16949 - 8.3.4", "IATF 16949 - 8.3.4.1", "IATF 16949 - 8.3.4.2", "IATF 16949 - 8.3.4.3", "IATF 16949 - 8.3.4.4", "IATF 16949 - 8.3.5", "IATF 16949 - 8.3.5.1", "IATF 16949 - 8.3.6", "IATF 16949 - 8.3.6.1", "IATF 16949 - 8.4", "IATF 16949 - 8.4.1", "IATF 16949 - 8.4.1.1", "IATF 16949 - 8.4.1.2", "IATF 16949 - 8.4.2", "IATF 16949 - 8.4.2.1", "IATF 16949 - 8.4.2.2", "IATF 16949 - 8.4.2.3", "IATF 16949 - 8.4.2.4", "IATF 16949 - 8.4.2.5", "IATF 16949 - 8.4.2.6", "IATF 16949 - 8.4.3", "IATF 16949 - 8.5", "IATF 16949 - 8.5.1", "IATF 16949 - 8.5.1.1", "IATF 16949 - 8.5.1.2", "IATF 16949 - 8.5.1.3", "IATF 16949 - 8.5.1.4", "IATF 16949 - 8.5.1.5", "IATF 16949 - 8.5.1.6", "IATF 16949 - 8.5.1.7", "IATF 16949 - 8.5.1.8", "IATF 16949 - 8.5.1.9", "IATF 16949 - 8.5.1.10", "IATF 16949 - 8.5.1.11", "IATF 16949 - 8.5.1.12", "IATF 16949 - 8.5.2", "IATF 16949 - 8.5.2.1", "IATF 16949 - 8.5.3", "IATF 16949 - 8.5.4", "IATF 16949 - 8.5.4.1", "IATF 16949 - 8.5.5", "IATF 16949 - 8.5.5.1", "IATF 16949 - 8.5.5.2", "IATF 16949 - 8.5.5.3", "IATF 16949 - 8.5.5.4", "IATF 16949 - 8.5.6", "IATF 16949 - 8.5.6.1", "IATF 16949 - 8.6", "IATF 16949 - 8.6.1", "IATF 16949 - 8.6.2", "IATF 16949 - 8.6.3", "IATF 16949 - 8.6.4", "IATF 16949 - 8.6.5", "IATF 16949 - 8.6.6", "IATF 16949 - 8.7", "IATF 16949 - 8.7.1", "IATF 16949 - 8.7.1.1", "IATF 16949 - 8.7.1.2", "IATF 16949 - 8.7.1.3", "IATF 16949 - 8.7.1.4",
    "IATF 16949 - 9.1", "IATF 16949 - 9.1.1", "IATF 16949 - 9.1.1.1", "IATF 16949 - 9.1.1.2", "IATF 16949 - 9.1.1.3", "IATF 16949 - 9.1.2", "IATF 16949 - 9.1.2.1", "IATF 16949 - 9.1.2.2", "IATF 16949 - 9.1.2.3", "IATF 16949 - 9.1.2.4", "IATF 16949 - 9.1.3", "IATF 16949 - 9.2", "IATF 16949 - 9.2.1", "IATF 16949 - 9.2.2", "IATF 16949 - 9.2.2.1", "IATF 16949 - 9.2.2.2", "IATF 16949 - 9.2.2.3", "IATF 16949 - 9.2.2.4", "IATF 16949 - 9.3", "IATF 16949 - 9.3.1", "IATF 16949 - 9.3.1.1", "IATF 16949 - 9.3.2", "IATF 16949 - 9.3.2.1", "IATF 16949 - 9.3.2.2",
    "IATF 16949 - 10.1", "IATF 16949 - 10.2", "IATF 16949 - 10.2.1", "IATF 16949 - 10.2.2", "IATF 16949 - 10.2.3", "IATF 16949 - 10.2.4", "IATF 16949 - 10.2.5", "IATF 16949 - 10.3", "IATF 16949 - 10.3.1",
]
# ISO 9001 / IATF 16949를 구분하지 않고 조항 번호만 통합하여 4.1부터 순서대로 표시
def _clause_number(label):
    m = re.search(r"(\d+(?:\.\d+)+|\d+)$", str(label))
    return m.group(1) if m else None

_clause_nums = set()
for _label in ISO_CLAUSES + IATF_CLAUSES:
    _n = _clause_number(_label)
    if _n:
        _clause_nums.add(_n)
CLAUSE_OPTIONS = sorted(_clause_nums, key=lambda x: tuple(int(v) for v in x.split(".")))
NC_CLAUSE_OPTIONS = ["--조항 선택--"] + CLAUSE_OPTIONS + ["직접 입력"]

# ──────────────────────────────────────────────
# 세션 초기화
# ──────────────────────────────────────────────
def _init():
    # 수정본 최초 실행 시 기존 세션의 테스트용 기본값을 초기화
    if st.session_state.get("app_version") != APP_VERSION:
        for _k in ["auditors", "process_df", "nc_list", "shift_cfgs", "client_name", "scope", "audit_date", "audit_type", "upload_sig", "open_shift_dialog", "generated"]:
            st.session_state.pop(_k, None)
        st.session_state.app_version = APP_VERSION
    if "auditors" not in st.session_state:
        st.session_state.auditors = [
            {"role":"--역할 선택--","eac":[],"kr":"--심사원 선택--","days":0.0},
        ]
    if "process_df" not in st.session_state:
        st.session_state.process_df = pd.DataFrame(columns=["프로세스명","분류","제조"])
    if "nc_list" not in st.session_state:
        st.session_state.nc_list = []
    if "shift_cfgs" not in st.session_state:
        st.session_state.shift_cfgs = {}
    st.session_state.setdefault("client_name","")
    st.session_state.setdefault("scope","")
    st.session_state.setdefault("audit_date", None)
    st.session_state.setdefault("audit_type","--심사 종류 선택--")
    st.session_state.setdefault("open_shift_dialog", False)
    st.session_state.setdefault("upload_sig", None)

_init()

# ──────────────────────────────────────────────
# 헤더
# ──────────────────────────────────────────────
st.title("📋 IATF 16949 CARA 양식 심사계획서 자동 생성기")
st.caption("공식 Word(.docx) 서식 변환 · 2교대 실제 근무시간대 동적 연동 · 제조 비중 31%~34% 통제")
st.divider()

# ══════════════════════════════════════════════
# 1️⃣  심사 기본 정보 및 심사원 구성
# ══════════════════════════════════════════════
st.header("1️⃣ 심사 기본 정보 및 심사원 구성")
st.caption("💡 **EAC Code**는 선임심사원에서 선택하시면 모든 심사원에게 동일하게 자동 일괄 적용됩니다. (필요 시 개별 변경도 가능)")

c1, c2 = st.columns(2)
with c1:
    if st.button("＋ 심사원 추가", use_container_width=True):
        lead_eac = st.session_state.auditors[0]["eac"] if st.session_state.auditors else []
        st.session_state.auditors.append({"role":"--역할 선택--","eac":list(lead_eac),"kr":"--심사원 선택--","days":0.0})
        st.rerun()
with c2:
    if st.button("－ 마지막 심사원 삭제", use_container_width=True) and len(st.session_state.auditors) > 1:
        st.session_state.auditors.pop()
        st.rerun()

for idx, aud in enumerate(st.session_state.auditors):
    label = "· 선임심사원 (Lead Auditor) 설정" if idx == 0 else f"· 심사원 #{idx+1} (Auditor) 설정"
    st.markdown(f"**{label}**")
    col_role, col_eac, col_q, col_kr, col_days = st.columns([2, 3, 0.2, 2, 1.5])
    with col_role:
        st.caption(f"역할 ({idx+1})")
        role_opts = ["--역할 선택--","Lead Auditor","Auditor"]
        aud["role"] = st.selectbox("역할", role_opts,
            index=role_opts.index(aud["role"]) if aud["role"] in role_opts else 0,
            key=f"role_{idx}", label_visibility="collapsed")
    with col_eac:
        st.caption(f"EAC Code 선택 ({idx+1})")
        sel_eac = st.multiselect("EAC", EAC_OPTIONS, default=aud["eac"],
            key=f"eac_{idx}", label_visibility="collapsed")
        if idx == 0 and sel_eac != aud["eac"]:
            for a in st.session_state.auditors:
                a["eac"] = list(sel_eac)
            aud["eac"] = sel_eac
            st.rerun()
        aud["eac"] = sel_eac
    with col_q:
        st.caption(" ")
        st.markdown("ℹ️", help="EAC Code는 심사원의 산업분야 자격코드입니다.")
    with col_kr:
        st.caption(f"국문 이름 선택 ({idx+1})")
        kr_idx = KR_NAMES.index(aud["kr"]) if aud["kr"] in KR_NAMES else 0
        chosen = st.selectbox("이름", KR_NAMES, index=kr_idx,
            key=f"kr_{idx}", label_visibility="collapsed")
        aud["kr"] = chosen
        if chosen != "--심사원 선택--":
            st.caption(f"🔗 매칭: {KR_TO_EN.get(chosen,'')}")
        else:
            st.caption("🔗 심사원을 선택해 주세요.")
    with col_days:
        st.caption(f"참여 일수 ({idx+1}, MD)")
        d_col1, d_col2, d_col3 = st.columns([1,2,1])
        with d_col1:
            if st.button("－", key=f"dminus_{idx}") and aud["days"] > 0.5:
                aud["days"] = round(aud["days"] - 0.5, 1); st.rerun()
        with d_col2:
            st.markdown(f"<div style='text-align:center;font-size:18px;font-weight:600;padding:4px'>{aud['days']}</div>", unsafe_allow_html=True)
        with d_col3:
            if st.button("＋", key=f"dplus_{idx}"):
                aud["days"] = round(aud["days"] + 0.5, 1); st.rerun()
    aud["en"] = KR_TO_EN.get(aud["kr"], aud.get("en",""))
    aud["display_name"] = aud["en"]
    aud["role_eac"] = aud["role"] + (" / " + ", ".join(aud["eac"]) if aud["eac"] else "")
    st.markdown("---")

# 고객 정보
st.markdown("&nbsp;")
ci1, ci2, ci3 = st.columns([2,3,2])
with ci1:
    st.session_state.client_name = st.text_input("고객명 (Client Name)", value=st.session_state.client_name)
with ci2:
    st.session_state.scope = st.text_area("인증범위 (Scope)", value=st.session_state.scope, height=80)
with ci3:
    audit_types = ["--심사 종류 선택--","사후심사 (Surveillance)","갱신심사 (Recertification)","최초심사 (Stage 2)","최초 1단계 (Stage 1)"]
    st.session_state.audit_type = st.selectbox("심사 종류 선택", audit_types,
        index=audit_types.index(st.session_state.audit_type) if st.session_state.audit_type in audit_types else 0)

ci4, ci5 = st.columns([2,2])
with ci4:
    st.session_state.audit_date = st.date_input("심사 시작일자", value=st.session_state.audit_date, format="YYYY-MM-DD")
with ci5:
    total_md = sum(a["days"] for a in st.session_state.auditors)
    d2c1, d2c2, d2c3 = st.columns([1,2,1])
    st.caption("총 심사 일수 (Total MD - 자동 합산)")
    st.markdown(f"<div style='font-size:22px;font-weight:700;padding:6px 0'>{total_md}</div>", unsafe_allow_html=True)

st.divider()

# ══════════════════════════════════════════════
# 2️⃣  심사 대상 프로세스 관리
# ══════════════════════════════════════════════
st.header("2️⃣ 심사 대상 프로세스 관리 (엑셀 업로드 및 기본정보 자동 추출)")
st.caption("💡 엑셀 업로드 시 A1~F1 메뉴 헤더 기준으로 **A열(조직명), B열(인증범위), C열(심사시작일)**은 기본 정보로 자동 입력되고, "
           "**D열(프로세스명), E열(프로세스 구분), F열(제조)**은 심사 프로세스로 자동 등록됩니다.")

up_col, dl_col = st.columns([4,1])
with up_col:
    uploaded = st.file_uploader("📥 프로세스 및 심사정보 엑셀 파일 업로드 (.xlsx, .xls, .csv)",
        type=["xlsx","xls","csv"], key="proc_upload")
with dl_col:
    st.caption(" ")
    # 신규 엑셀 양식 다운로드
    tmpl = pd.DataFrame({
        "조직명(A)":["예: (주)서원인텍"],"인증범위(B)":["예: 고무성형부품의 제조"],
        "심사시작일(C)":["2026-10-06"],"프로세스명(D)":["Management review P"],
        "프로세스구분(E)":["MP"],"제조여부(F)":["False"]
    })
    buf = io.BytesIO(); tmpl.to_excel(buf, index=False); buf.seek(0)
    st.download_button("📄 신규 엑셀 양식 다운로드", buf,
        file_name="CARA_프로세스_양식.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True)

if uploaded is not None:
    sig = (uploaded.name, uploaded.size)
    if sig != st.session_state.upload_sig:
        try:
            raw = pd.read_excel(uploaded, header=None) if not uploaded.name.endswith(".csv") else pd.read_csv(uploaded, header=None)
            # 헤더 행 탐색
            found = False
            for i in range(min(len(raw), 10)):
                row = [str(v).strip().lower() if pd.notna(v) else "" for v in raw.iloc[i]]
                if any("프로세스" in r or "process" in r for r in row):
                    cols = [str(v).strip() if pd.notna(v) else "" for v in raw.iloc[i]]
                    body = raw.iloc[i+1:].reset_index(drop=True)
                    body.columns = range(len(body.columns))
                    # A~F 고정 위치
                    org = str(body.iloc[0,0]).strip() if len(body.columns)>0 else ""
                    scope_val = str(body.iloc[0,1]).strip() if len(body.columns)>1 else ""
                    date_val = body.iloc[0,2] if len(body.columns)>2 else None
                    rows_proc = []
                    for _, r in body.iterrows():
                        nm = str(r.get(3,"")).strip() if len(body.columns)>3 else ""
                        cl = str(r.get(4,"")).strip() if len(body.columns)>4 else ""
                        mf_raw = r.get(5, False) if len(body.columns)>5 else False
                        if not nm or nm=="nan": continue
                        mf = str(mf_raw).strip().lower() in ("true","1","yes","예","o","제조","✔","✓","t","ture")
                        rows_proc.append({"프로세스명":nm,"분류":cl,"제조":mf})
                    if rows_proc:
                        st.session_state.process_df = pd.DataFrame(rows_proc)
                        if org and org != "nan": st.session_state.client_name = org
                        if scope_val and scope_val != "nan": st.session_state.scope = scope_val
                        if date_val is not None:
                            try:
                                st.session_state.audit_date = pd.to_datetime(date_val).date()
                            except: pass
                        st.session_state.nc_list = []
                        st.session_state.shift_cfgs = {}
                        st.session_state.open_shift_dialog = True
                        st.session_state.upload_sig = sig
                        n_proc = len(rows_proc)
                        st.success(f"✅ 엑셀 파일에서 기본 정보 (조직명: **{org}**, 인증범위: **{scope_val[:20]}...**, "
                                   f"심사시작일: **{st.session_state.audit_date}**) 및 프로세스 **{n_proc}개**를 성공적으로 불러왔습니다!")
                        found = True; break
            if not found:
                st.error("❌ A~F열 헤더(조직명~제조)를 찾지 못했습니다. 양식을 확인해 주세요.")
        except Exception as e:
            st.error(f"❌ 파일 읽기 오류: {e}")

# 프로세스 테이블
df = st.session_state.process_df.copy()
if df.empty:
    df = pd.DataFrame(columns=["프로세스명","분류","제조"])
st.session_state.process_df = st.data_editor(df, num_rows="dynamic", use_container_width=True,
    key="proc_editor",
    column_config={
        "프로세스명": st.column_config.TextColumn("프로세스명", help="엑셀 업로드 전에는 비어 있습니다."),
        "분류": st.column_config.TextColumn("프로세스 구분"),
        "제조": st.column_config.CheckboxColumn("제조 프로세스 여부 (30~35% 대상)", default=False)
    })

# 교대조 설정 (제조 프로세스별)
mfg_procs = st.session_state.process_df[st.session_state.process_df["제조"]==True]["프로세스명"].dropna().astype(str).tolist()

@st.dialog("🏭 제조 프로세스 교대조 운영 설정", width="large")
def show_shift_settings_dialog(processes):
    st.caption(f"업로드된 제조 프로세스 **{len(processes)}개**의 교대조 운영을 설정합니다.")
    st.info("이 창은 프로세스 엑셀을 새로 업로드하면 자동으로 열립니다.")
    for k, proc in enumerate(processes):
        st.markdown(f"**⚙️ [{k+1}/{len(processes)}] {proc} 교대조 설정**")
        cfg = st.session_state.shift_cfgs.get(proc, {"type":"1교대","s1":"08:30-17:30","s2":"17:30-02:30","s2_audit":"1.0","s2_slot":"17:00 - 18:00"})
        shift_types = ["1교대 (단일조, 교대근무 없음)","2교대 (주/야간)","3교대 (3조 3교대/3조 2교대)"]
        type_names = ["1교대","2교대","3교대"]
        shift_type = st.radio(f"교대조 운영 형태 ({proc})", shift_types,
            index=type_names.index(cfg["type"]) if cfg.get("type") in type_names else 0,
            key=f"dialog_shift_type_{k}", horizontal=True)
        cfg["type"] = shift_type.split(" ")[0]
        cfg["s1"] = st.text_input(f"1교대(주간) 근무시간 ({proc})", value=cfg.get("s1","08:30-17:30"), key=f"dialog_s1_{k}")
        if "2교대" in shift_type:
            c1s, c2s = st.columns(2)
            with c1s: cfg["s2"] = st.text_input("2교대(야간) 근무시간", value=cfg.get("s2","17:30-02:30"), key=f"dialog_s2_{k}")
            with c2s: cfg["s2_slot"] = st.selectbox("야간 심사 시작 시간", ["17:00 - 18:00","18:00 - 19:00","20:00 - 21:00"], index=0, key=f"dialog_s2slot_{k}")
            cfg["s2_audit"] = st.slider("야간 심사 배정 시간(h)", 0.5, 2.0, float(cfg.get("s2_audit",1.0)), 0.5, key=f"dialog_s2aud_{k}")
        st.session_state.shift_cfgs[proc] = cfg
        if k < len(processes)-1: st.markdown("---")
    if st.button("확인 / 설정 저장", type="primary", use_container_width=True):
        st.rerun()

if mfg_procs and st.session_state.get("open_shift_dialog", False):
    st.session_state.open_shift_dialog = False
    show_shift_settings_dialog(mfg_procs)

if mfg_procs:
    with st.expander(f"🏭 제조 프로세스 교대조(Shift) 운영 설정 다시 열기", expanded=False):
        st.caption("필요할 때 교대조 설정을 다시 확인하거나 수정할 수 있습니다.")
        if st.button("교대조 설정 열기", key="open_shift_again", use_container_width=True):
            show_shift_settings_dialog(mfg_procs)

st.divider()

# ══════════════════════════════════════════════
# 3️⃣  이전 부적합 시정조치 검증 관리
# ══════════════════════════════════════════════
st.header("3️⃣ 이전 부적합 시정조치 검증 관리 (추가 심사시간)")

proc_names = st.session_state.process_df["프로세스명"].tolist()
with st.expander("➕ 부적합 검증 항목 추가", expanded=False):
    with st.form("nc_add_form", clear_on_submit=True):
        fc1, fc2, fc3, fc4 = st.columns([3,3,1.5,1])
        with fc1: nc_proc = st.selectbox("대상 프로세스", proc_names) if proc_names else st.selectbox("대상 프로세스", ["--프로세스 없음--"])
        with fc2:
            nc_clause = st.selectbox("조항", NC_CLAUSE_OPTIONS, key="nc_clause_select")
            if nc_clause == "직접 입력":
                nc_clause = st.text_input("조항 직접 입력", placeholder="예: 8.5.1.1")
        with fc3: nc_hours = st.number_input("추가 시간(h)", min_value=0.5, max_value=4.0, value=0.5, step=0.5)
        with fc4:
            st.caption(" "); st.caption(" ")
            submitted = st.form_submit_button("추가", use_container_width=True)
        if submitted and proc_names and nc_proc != "--프로세스 없음--" and nc_clause and nc_clause != "--조항 선택--":
            st.session_state.nc_list.append({"proc_name":nc_proc,"clauses":nc_clause,"extra_hours":nc_hours})
            st.rerun()

for i, nc in enumerate(st.session_state.nc_list):
    nc1, nc2, nc3, nc4 = st.columns([3,3,1.5,1])
    with nc1: st.markdown(f"**대상:** {nc['proc_name']}")
    with nc2: st.markdown(f"**조항:** {nc['clauses']}")
    with nc3: st.markdown(f"**추가 시간: <span style='color:#c0392b;font-weight:700'>{nc['extra_hours']} hr</span>**", unsafe_allow_html=True)
    with nc4:
        if st.button("🗑 삭제", key=f"del_nc_{i}"):
            st.session_state.nc_list.pop(i); st.rerun()

st.divider()

def clean_process_name(name):
    return re.sub(r"^\[(COP|MP|SP)\]\s*", "", str(name)).strip()

def assign_processes_intelligently(df_proc, aud_index, total_auditors):
    temp = df_proc.copy()
    if total_auditors <= 1:
        return temp
    
    # Lead Auditor: 경영, 제조(엑셀 제조=True 포함), 설비/보전
    lead_kw = ["경영", "리더십", "기획", "내부심사", "리스크", "생산", "제조", "설비", "금형", "보전"]
    team_kw = ["영업", "계약", "설계", "개발", "구매", "자재", "공급", "품질", "검사", "시험", "출하", "인도", "고객", "부적합", "개선"]
    is_mfg = temp["제조프로세스여부"] == True
    lead_mask = temp["프로세스명"].str.contains("|".join(lead_kw), case=False, na=False) | is_mfg
    team_mask = temp["프로세스명"].str.contains("|".join(team_kw), case=False, na=False) & ~lead_mask
    unmatched = ~(lead_mask | team_mask)   # 키워드에 안 걸린 프로세스(엑셀 업로드 시 흔함)도 누락 없이 Team에 배정
    if aud_index == 0:
        return temp[lead_mask]
    return temp[team_mask | unmatched]

DAY_START = 9 * 60       # 09:00
LUNCH_S, LUNCH_E = 12 * 60, 13 * 60
DAY_END = 18 * 60        # 09:00~18:00, 점심 제외 = 정확히 8.0h = 1 MD


def _hm(m):
    return f"{m // 60:02d}:{m % 60:02d}"


def _to_min(s):
    h, m = str(s).strip()[:5].split(":")
    return int(h) * 60 + int(m)


def _split_units(total_h, n):
    """total_h 시간을 0.5h 단위로 n개에 균등 분배 (합계 정확히 일치)."""
    if n <= 0:
        return []
    units = int(round(total_h * 2))
    base, rem = divmod(units, n)
    return [(base + (1 if i < rem else 0)) / 2 for i in range(n)]


class _Timeline:
    """커서를 이동하며 실제 시각을 계산 → 표시 시간과 Duration이 항상 일치."""

    def __init__(self):
        self.day = 1
        self.t = DAY_START
        self.rows = []
        self.reserved = {}   # day -> [(start, end)] 교대조 등 고정 슬롯

    def _skip(self):
        moved = True
        while moved:
            moved = False
            if LUNCH_S <= self.t < LUNCH_E:
                self.t, moved = LUNCH_E, True
            for s, e in self.reserved.get(self.day, []):
                if s <= self.t < e:
                    self.t, moved = e, True

    def _stop(self):
        stops = [DAY_END]
        if self.t < LUNCH_S:
            stops.append(LUNCH_S)
        stops += [s for s, _ in self.reserved.get(self.day, []) if s > self.t]
        return min(stops)

    def _next_day(self):
        self.day += 1
        self.t = DAY_START
        self.add("Daily meeting 일일 회의", 0.5, atomic=True)

    def add(self, label, hours, atomic=False, extra=0.0, continuation_label=None):
        """시간을 배치합니다. continuation_label을 주면 (계속) 행에는 별도 라벨을 사용합니다."""
        remaining = int(round(hours * 60))
        first = True
        while remaining > 0:
            self._skip()
            if self.t >= DAY_END:
                self._next_day()
                continue
            stop = self._stop()
            if atomic and stop - self.t < remaining:   # 회의는 쪼개지 않음
                if stop >= DAY_END:
                    self._next_day()
                else:
                    self.t = stop
                continue
            chunk = min(remaining, stop - self.t)
            cont = continuation_label if continuation_label is not None else f"{label} (계속)"
            self.rows.append({
                "day": self.day, "start": self.t, "end": self.t + chunk,
                "process": label if first else cont,
                "dur_real": chunk / 60.0,
                "dur_extra": extra if first else 0.0,
            })
            self.t += chunk
            remaining -= chunk
            first = False

    def free_slot(self, day, start, minutes):
        """이미 예약된 슬롯과 겹치면 그 뒤로 미룸 (제조 프로세스가 여러 개일 때)."""
        moved = True
        while moved:
            moved = False
            for s, e in self.reserved.get(day, []):
                if start < e and start + minutes > s:
                    start, moved = e, True
        return start

    def add_fixed(self, day, start, hours, label):
        self.rows.append({"day": day, "start": start, "end": start + int(round(hours * 60)),
                          "process": label, "dur_real": hours, "dur_extra": 0.0})


def _sequence_priority(row):
    name, is_mfg = row["프로세스명"], row["제조프로세스여부"]
    if is_mfg: return 80
    if any(k in name for k in ["검사", "품질", "부적합"]): return 70
    if any(k in name for k in ["경영", "리더십", "기획"]): return 10
    if any(k in name for k in ["영업", "계약"]): return 20
    if any(k in name for k in ["설계", "개발"]): return 30
    if any(k in name for k in ["구매", "자재", "공급"]): return 40
    return 50


def _build(records, target, nd, nc_df, shift_cfg):
    fixed = 1.5 + 0.5 * (nd - 1) + 1.5            # 시작회의+순회 / 일일회의 / 팀회의+종료회의
    n_m = sum(1 for r in records if r["제조프로세스여부"])
    n_o = len(records) - n_m
    P = max(0.5 * len(records), target - fixed)    # 프로세스 심사에 쓸 총 시간

    # 제조 비중 31~34% (구간 내 중간값, 0.5h 단위)
    if n_m:
        lo = math.ceil(target * MIN_MFG_RATIO * 2) / 2
        hi = max(math.floor(target * MAX_MFG_RATIO * 2) / 2, lo)
        mfg_h = min(max(round(target * 0.325 * 2) / 2, lo), hi)
        mfg_h = max(0.5 * n_m, min(mfg_h, P - 0.5 * n_o)) if n_o else P
    else:
        mfg_h = 0.0
    m_durs = iter(_split_units(mfg_h, n_m))
    o_durs = iter(_split_units(P - mfg_h, n_o))

    tl = _Timeline()
    tl.add("Opening meeting 시작회의", 1.0, atomic=True)
    tl.add("Line tour & CEO meeting 현장순회 및 경영자 면담", 0.5, atomic=True)

    has_shift = shift_cfg.get("has_shifts")
    for r in records:
        dur = next(m_durs) if r["제조프로세스여부"] else next(o_durs)
        base_name = re_clean(r["프로세스명"])
        name = base_name
        extra = 0.0
        nc_note = ""
        if nc_df is not None and not nc_df.empty:
            hit = nc_df[nc_df["proc_name"] == r["프로세스명"]]
            if not hit.empty:
                extra = float(hit.iloc[0]["extra_hours"])
                clause_text = str(hit.iloc[0]["clauses"]).strip()
                clause_text = _clause_number(clause_text) or clause_text
                nc_note = f"Previous audit NC effective verification : {clause_text} / {extra:.1f} hr"
                name = f"{base_name}\n{nc_note}"

        is_shift_target = (has_shift and r["제조프로세스여부"] and
                           r["프로세스명"] in shift_cfg.get("target_mfg_procs", []))
        if is_shift_target and dur > 0.5:
            s2 = min(float(shift_cfg.get("s2_dur", 1.0)), dur - 0.5)
            slot = tl.free_slot(tl.day, _to_min(shift_cfg.get("s2_audit_slot", "17:00")), int(s2 * 60))
            # 이전 부적합 검증 문구/시간은 해당 프로세스의 최초 심사행에만 1회 표시
            s1_base = f"1 shift ({shift_cfg.get('s1_hours', '')}) {base_name}"
            s1_label = f"1 shift ({shift_cfg.get('s1_hours', '')}) {name}" if nc_note else s1_base
            s2_label = f"2 shift ({shift_cfg.get('s2_hours', '')}) {base_name} 야간 교대조 심사"
            tl._skip()
            if tl.t < slot:   # 아직 슬롯 전 → 미리 예약해서 주간 심사가 겹치지 않게 함
                tl.reserved.setdefault(tl.day, []).append((slot, slot + int(s2 * 60)))
                tl.add_fixed(tl.day, slot, s2, s2_label)
                tl.add(s1_label, dur - s2, extra=extra, continuation_label=f"{s1_base} (계속)")
            else:             # 이미 늦었으면 주간 심사 후 바로 이어서 배치
                tl.add(s1_label, dur - s2, extra=extra, continuation_label=f"{s1_base} (계속)")
                start = tl.free_slot(tl.day, max(slot, tl.t), int(s2 * 60))
                tl.reserved.setdefault(tl.day, []).append((start, start + int(s2 * 60)))
                tl.add_fixed(tl.day, start, s2, s2_label)
        else:
            tl.add(name, dur, extra=extra, continuation_label=f"{base_name} (계속)")

    # 야간 교대조 심사가 같은 날 남아 있으면 팀회의/종료회의는 그 이후로
    late = [e for _, e in tl.reserved.get(tl.day, [])]
    if late and max(late) > tl.t:
        t0 = max(late)   # 같은 날 야간 심사 직후에 직접 배치 (18:00 이후 허용)
        tl.add_fixed(tl.day, t0, 0.5, "Auditors meeting 심사팀 회의")
        tl.add_fixed(tl.day, t0 + 30, 1.0, "Closing meeting 종료회의")
    else:
        tl.add("Auditors meeting 심사팀 회의", 0.5, atomic=True)
        tl.add("Closing meeting 종료회의", 1.0, atomic=True)
    return tl.rows, tl.day


def re_clean(name):
    import re
    return re.sub(r"^\[(COP|MP|SP)\]\s*", "", str(name)).strip()


def generate_cara_schedule_for_auditor(df_proc, nc_df, auditor_info, aud_index, total_auditors, shift_cfg):
    att_days = float(auditor_info.get("days_attending", 1.0))
    target = round(att_days * HOURS_PER_MD, 1)

    assigned = assign_processes_intelligently(df_proc, aud_index, total_auditors)
    if assigned.empty:
        assigned = df_proc.head(2)
    assigned = assigned.copy()
    assigned["seq_pri"] = assigned.apply(_sequence_priority, axis=1)
    records = assigned.sort_values("seq_pri", kind="stable").to_dict("records")

    nd = max(1, math.ceil(att_days))
    for _ in range(3):                      # 일수가 넘치면 일일회의 시간을 반영해 재계산
        rows, used = _build(records, target, nd, nc_df, shift_cfg)
        if used == nd:
            break
        nd = used

    rows.sort(key=lambda x: (x["day"], x["start"]))
    out = list(rows)
    for d in sorted({r["day"] for r in rows}):                 # 점심 행 자동 삽입
        dr = [r for r in rows if r["day"] == d]
        if min(r["start"] for r in dr) < LUNCH_S and max(r["end"] for r in dr) > LUNCH_E:
            out.append({"day": d, "start": LUNCH_S, "end": LUNCH_E, "process": "Lunch 중식",
                        "dur_real": 0.0, "dur_extra": 0.0})
    out.sort(key=lambda x: (x["day"], x["start"]))
    for r in out:
        r["time"] = f"{_hm(r['start'])} - {_hm(r['end'])}"
    df = pd.DataFrame(out)
    return df[["day", "time", "process", "dur_real", "dur_extra"]]

# ----------------------------------------------------
# 4. 공식 CARA Word(.docx) 생성 함수 (10열 표준 그리드)
# ----------------------------------------------------
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>'))

def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '</w:tblBorders>'
    )
    tblPr.append(borders)

def format_cell(cell, text, fill=None, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=8.0, font_name="Calibri", font_color=None):
    p = cell.paragraphs[0]
    p.text = ""
    p.alignment = align
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.05
    
    lines = str(text).split("\n")
    for i, line in enumerate(lines):
        if i > 0:
            p = cell.add_paragraph()
            p.alignment = align
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.line_spacing = 1.05
        run = p.add_run(line)
        run.bold = bold
        run.font.name = font_name
        run.font.size = Pt(font_size)
        if font_color:
            run.font.color.rgb = RGBColor.from_string(font_color)
    
    if fill:
        set_cell_background(cell, fill)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_margins(cell)

def format_process_cell(cell, text, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=8.0):
    """프로세스명은 검정, 이전 부적합 검증 안내(영문~시간)는 붉은색으로 표시."""
    p = cell.paragraphs[0]
    p.text = ""
    p.alignment = align
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.05
    for i, line in enumerate(str(text).split("\n")):
        if i > 0:
            p = cell.add_paragraph()
            p.alignment = align
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.line_spacing = 1.05
        run = p.add_run(line)
        run.font.name = "Calibri"
        run.font.size = Pt(font_size)
        if i > 0 and "Previous audit NC effective verification" in line:
            run.bold = True
            run.font.color.rgb = RGBColor.from_string("C0392B")
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_margins(cell)

def create_cara_word_document(client_name, scope_text, start_date, auditors_info, all_schedules):
    doc = Document()
    
    # A4 세로 (Portrait: 11906 x 16838 dxa), 여백 0.5인치 (720 dxa)
    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    total_days = max(len(sched["day"].unique()) for sched in all_schedules) if all_schedules else 1
    is_paired = len(auditors_info) >= 2

    for day_idx in range(1, total_days + 1):
        if day_idx > 1:
            doc.add_page_break()

        # 1. 상단 타이틀
        p1 = doc.add_paragraph()
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(2)
        r1 = p1.add_run("AUDIT PLAN 심사계획서")
        r1.bold = True
        r1.font.name = "Calibri"
        r1.font.size = Pt(16)

        # 2. 설명문
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(1)
        r2 = p2.add_run("To add up audit duration totals, click anywhere in the audit plan, press Ctrl+a, then select F9")
        r2.font.name = "Arial"
        r2.font.size = Pt(9)

        p3 = doc.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p3.paragraph_format.space_before = Pt(0)
        p3.paragraph_format.space_after = Pt(6)
        r3 = p3.add_run("심사기간 총계를 추가하려면 심사계획서의 아무 곳이나 클릭하고 Ctrl+a를 누른 다음 F9를 클릭하세요.")
        r3.font.name = "Malgun Gothic"
        r3.font.size = Pt(9)

        curr_audit_date = (start_date + timedelta(days=day_idx - 1)).strftime("%d/%m/%Y")

        sched1 = all_schedules[0][all_schedules[0]["day"] == day_idx].to_dict("records") if len(all_schedules) > 0 else []
        sched2 = all_schedules[1][all_schedules[1]["day"] == day_idx].to_dict("records") if len(all_schedules) > 1 else []
        max_items = max(len(sched1), len(sched2)) if is_paired else len(sched1)

        num_rows = 4 + max_items + 4
        table = doc.add_table(rows=num_rows, cols=10)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        set_table_borders(table)

        col_widths_dxa = [706, 567, 2977, 568, 567, 709, 567, 2976, 568, 568]
        tblGrid_xml = f'<w:tblGrid {nsdecls("w")}>' + ''.join([f'<w:gridCol w:w="{w}"/>' for w in col_widths_dxa]) + '</w:tblGrid>'
        table._tbl.insert(1, parse_xml(tblGrid_xml))

        # Row 0: Date
        table.cell(0, 0).merge(table.cell(0, 1))
        format_cell(table.cell(0, 0), "Date   심사일자", fill="FDE9D9", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
        table.cell(0, 2).merge(table.cell(0, 4))
        format_cell(table.cell(0, 2), curr_audit_date, align=WD_ALIGN_PARAGRAPH.CENTER)

        table.cell(0, 5).merge(table.cell(0, 6))
        format_cell(table.cell(0, 5), "Date   심사일자", fill="FDE9D9", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
        table.cell(0, 7).merge(table.cell(0, 9))
        format_cell(table.cell(0, 7), curr_audit_date if is_paired else "-", align=WD_ALIGN_PARAGRAPH.CENTER)

        # Row 1: Auditor Name
        aud1_text = f"{auditors_info[0]['display_name']}\n{auditors_info[0]['role']}" if len(auditors_info) > 0 else ""
        aud2_text = f"{auditors_info[1]['display_name']}\n{auditors_info[1]['role']}" if is_paired else "-"

        table.cell(1, 0).merge(table.cell(1, 1))
        format_cell(table.cell(1, 0), "Auditor Name/Role/  EA Code\n심사원 이름/역할/  EA Code", fill="FDE9D9", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
        table.cell(1, 2).merge(table.cell(1, 4))
        format_cell(table.cell(1, 2), aud1_text, align=WD_ALIGN_PARAGRAPH.LEFT)

        table.cell(1, 5).merge(table.cell(1, 6))
        format_cell(table.cell(1, 5), "Auditor Name/Role/  EA Code\n심사원 이름/역할/EA Code", fill="FDE9D9", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
        table.cell(1, 7).merge(table.cell(1, 9))
        format_cell(table.cell(1, 7), aud2_text, align=WD_ALIGN_PARAGRAPH.LEFT)

        # Row 2: Scope
        table.cell(2, 0).merge(table.cell(2, 1))
        format_cell(table.cell(2, 0), "Scope 인증범위", fill="FDE9D9", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
        table.cell(2, 2).merge(table.cell(2, 9))
        format_cell(table.cell(2, 2), scope_text, align=WD_ALIGN_PARAGRAPH.LEFT)

        # Row 3: Header
        format_cell(table.cell(3, 0), "TIME\n시간", fill="FDE9D9", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        table.cell(3, 1).merge(table.cell(3, 2))
        format_cell(table.cell(3, 1), "Location/Process\n위치 / 프로세스", fill="FDE9D9", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        table.cell(3, 3).merge(table.cell(3, 4))
        format_cell(table.cell(3, 3), "Duration\n기간", fill="FDE9D9", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

        format_cell(table.cell(3, 5), "TIME\n시간", fill="FDE9D9", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        table.cell(3, 6).merge(table.cell(3, 7))
        format_cell(table.cell(3, 6), "Location/Process\n위치 / 프로세스", fill="FDE9D9", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        table.cell(3, 8).merge(table.cell(3, 9))
        format_cell(table.cell(3, 8), "Duration\n기간", fill="FDE9D9", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

        # Data Rows
        for i in range(max_items):
            r_idx = 4 + i
            if i < len(sched1):
                item1 = sched1[i]
                t_str1 = item1["time"].replace(" - ", "\n").replace(" ~ ", "\n")
                format_cell(table.cell(r_idx, 0), t_str1, align=WD_ALIGN_PARAGRAPH.CENTER)
                table.cell(r_idx, 1).merge(table.cell(r_idx, 2))
                format_process_cell(table.cell(r_idx, 1), item1["process"], align=WD_ALIGN_PARAGRAPH.LEFT)
                format_cell(table.cell(r_idx, 3), f"{item1['dur_real']:.1f}" if item1['dur_real'] > 0 else "0", align=WD_ALIGN_PARAGRAPH.CENTER)
                format_cell(table.cell(r_idx, 4), f"{item1['dur_extra']:.1f}" if item1['dur_extra'] > 0 else "0", align=WD_ALIGN_PARAGRAPH.CENTER, font_color="C0392B" if item1['dur_extra'] > 0 else None, bold=item1['dur_extra'] > 0)
            else:
                table.cell(r_idx, 1).merge(table.cell(r_idx, 2))

            if is_paired and i < len(sched2):
                item2 = sched2[i]
                t_str2 = item2["time"].replace(" - ", "\n").replace(" ~ ", "\n")
                format_cell(table.cell(r_idx, 5), t_str2, align=WD_ALIGN_PARAGRAPH.CENTER)
                table.cell(r_idx, 6).merge(table.cell(r_idx, 7))
                format_process_cell(table.cell(r_idx, 6), item2["process"], align=WD_ALIGN_PARAGRAPH.LEFT)
                format_cell(table.cell(r_idx, 8), f"{item2['dur_real']:.1f}" if item2['dur_real'] > 0 else "0", align=WD_ALIGN_PARAGRAPH.CENTER)
                format_cell(table.cell(r_idx, 9), f"{item2['dur_extra']:.1f}" if item2['dur_extra'] > 0 else "0", align=WD_ALIGN_PARAGRAPH.CENTER, font_color="C0392B" if item2['dur_extra'] > 0 else None, bold=item2['dur_extra'] > 0)
            else:
                table.cell(r_idx, 6).merge(table.cell(r_idx, 7))

        tot_act1 = sum(s["dur_real"] for s in sched1)
        tot_add1 = sum(s["dur_extra"] for s in sched1)
        tot_dur1 = round(tot_act1 + tot_add1, 1)

        tot_act2 = sum(s["dur_real"] for s in sched2) if is_paired else 0.0
        tot_add2 = sum(s["dur_extra"] for s in sched2) if is_paired else 0.0
        tot_dur2 = round(tot_act2 + tot_add2, 1)

        # Footer Row 1: Total Daily Audit Time
        r_time = 4 + max_items
        table.cell(r_time, 0).merge(table.cell(r_time, 2))
        format_cell(table.cell(r_time, 0), "Total Daily Audit Time (hours)\n일일 총 심사시간(hours)", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
        format_cell(table.cell(r_time, 3), f"{tot_act1:.1f}", fill="FAE2D5", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        format_cell(table.cell(r_time, 4), f"{tot_add1:.1f}", fill="FAE2D5", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

        table.cell(r_time, 5).merge(table.cell(r_time, 7))
        format_cell(table.cell(r_time, 5), "Total Daily Audit Time (hours)\n일일 총 심사시간(hours)", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
        format_cell(table.cell(r_time, 8), f"{tot_act2:.1f}" if is_paired else "-", fill="FAE2D5" if is_paired else None, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        format_cell(table.cell(r_time, 9), f"{tot_add2:.1f}" if is_paired else "-", fill="FAE2D5" if is_paired else None, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

        # Footer Row 2: Total Daily Additional Audit Time
        r_add = r_time + 1
        table.cell(r_add, 0).merge(table.cell(r_add, 4))
        format_cell(table.cell(r_add, 0), "Total Daily Additional Audit Time (hours)\n일일 총 추가적 심사시간(hours)", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
        table.cell(r_add, 5).merge(table.cell(r_add, 9))
        format_cell(table.cell(r_add, 5), "Total Daily Additional Audit Time (hours)\n일일 총 추가적 심사시간(hours)", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)

        # Footer Row 3: Total Daily Audit Duration
        r_dur = r_time + 2
        table.cell(r_dur, 0).merge(table.cell(r_dur, 3))
        format_cell(table.cell(r_dur, 0), "Total Daily Audit Duration (hours)\n일일 총 심사기간(hours)", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
        format_cell(table.cell(r_dur, 4), f"{tot_dur1:.1f}", fill="FAE2D5", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

        table.cell(r_dur, 5).merge(table.cell(r_dur, 8))
        format_cell(table.cell(r_dur, 5), "Total Daily Audit Duration (hours)\n일일 총 심사기간(hours)", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
        format_cell(table.cell(r_dur, 9), f"{tot_dur2:.1f}" if is_paired else "-", fill="FAE2D5" if is_paired else None, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

        # Footer Row 4: Disclaimer
        r_disc = r_time + 3
        table.cell(r_disc, 0).merge(table.cell(r_disc, 9))
        disc_text = (
            "Audit trails will be developed based upon identified risk throughout the audit and as such timings and content may be subject to change\n"
            "심사 진행은 심사 동안 파악된 리스크에 따라 변경될 것이며, 이에 따라 시기 및 내용이 변경될 수 있다."
        )
        format_cell(table.cell(r_disc, 0), disc_text, fill="FDE9D9", bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=7.5)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ══════════════════════════════════════════════
# 생성 버튼
# ══════════════════════════════════════════════
if st.button("🚀 CARA 양식 심사계획서 생성", type="primary", use_container_width=True):
    errs = []
    if not st.session_state.client_name.strip(): errs.append("고객명")
    if not st.session_state.scope.strip(): errs.append("인증범위")
    if st.session_state.audit_date is None: errs.append("심사 시작일자")
    if st.session_state.audit_type == "--심사 종류 선택--": errs.append("심사 종류 선택")
    if any(a["role"]=="--역할 선택--" for a in st.session_state.auditors): errs.append("심사원 역할 선택 미완료")
    if any(a["kr"]=="--심사원 선택--" for a in st.session_state.auditors): errs.append("심사원 선택 미완료")
    if errs:
        st.error(f"❌ 필수 항목 누락: **{', '.join(errs)}**")
        st.stop()

    # shift_config 통합 (기존 스케줄러 호환)
    mfg_list = st.session_state.process_df[st.session_state.process_df["제조"]==True]["프로세스명"].tolist()
    # 첫 번째 2교대 설정을 대표로 사용 (기존 스케줄러 API)
    shift_config = {"has_shifts": False, "target_mfg_procs": mfg_list, "s2_dur": 1.0, "s2_audit_slot": "17:00 - 18:00", "s1_hours": "", "s2_hours": ""}
    for proc, cfg in st.session_state.shift_cfgs.items():
        if cfg.get("type") == "2교대":
            shift_config["has_shifts"] = True
            shift_config["s1_hours"] = cfg.get("s1","")
            shift_config["s2_hours"] = cfg.get("s2","")
            shift_config["s2_dur"] = float(cfg.get("s2_audit",1.0))
            shift_config["s2_audit_slot"] = cfg.get("s2_slot","17:00 - 18:00")
            break

    # process_df 컬럼명 호환 (기존 스케줄러는 "제조프로세스여부" 사용)
    proc_df = st.session_state.process_df.copy()
    if "제조" in proc_df.columns and "제조프로세스여부" not in proc_df.columns:
        proc_df["제조프로세스여부"] = proc_df["제조"]

    nc_df = pd.DataFrame(st.session_state.nc_list) if st.session_state.nc_list else pd.DataFrame(columns=["proc_name","clauses","extra_hours"])
    parsed_auditors = []
    for a in st.session_state.auditors:
        parsed_auditors.append({
            "display_name": a["en"], "role": a["role_eac"],
            "days_attending": a["days"], "kr": a["kr"]
        })

    all_schedules = []
    for i, aud in enumerate(parsed_auditors):
        s = generate_cara_schedule_for_auditor(proc_df, nc_df, aud, i, len(parsed_auditors), shift_config)
        all_schedules.append(s)

    st.session_state.generated = {"schedules": all_schedules, "auditors": parsed_auditors,
        "shift_config": shift_config, "proc_df": proc_df, "nc_df": nc_df}
    st.success("✅ 심사계획서 생성이 완료되었습니다! 아래에서 요청하신 공식 양식 그대로 확인 및 Word(.docx) 다운로드가 가능합니다.")

# ══════════════════════════════════════════════
# 결과 표시
# ══════════════════════════════════════════════
if "generated" not in st.session_state:
    st.stop()

gen = st.session_state.generated
all_schedules = gen["schedules"]
parsed_auditors = gen["auditors"]
audit_date_input = st.session_state.audit_date
client_name = st.session_state.client_name
scope = st.session_state.scope
shift_config = gen["shift_config"]

total_days = max(len(s["day"].unique()) for s in all_schedules) if all_schedules else 1
is_paired = len(parsed_auditors) >= 2

for day_idx in range(1, total_days + 1):
    curr_date = (datetime.combine(audit_date_input, datetime.min.time()) + timedelta(days=day_idx-1)).strftime("%d/%m/%Y")
    sched1 = all_schedules[0][all_schedules[0]["day"]==day_idx].to_dict("records") if all_schedules else []
    sched2 = all_schedules[1][all_schedules[1]["day"]==day_idx].to_dict("records") if is_paired else []
    max_items = max(len(sched1), len(sched2)) if is_paired else len(sched1)
    tot_act1 = sum(s["dur_real"] for s in sched1)
    tot_add1 = sum(s["dur_extra"] for s in sched1)
    tot_dur1 = round(tot_act1 + tot_add1, 1)
    tot_act2 = sum(s["dur_real"] for s in sched2) if is_paired else 0.0
    tot_add2 = sum(s["dur_extra"] for s in sched2) if is_paired else 0.0
    tot_dur2 = round(tot_act2 + tot_add2, 1)
    aud1_html = f"{parsed_auditors[0]['display_name']}<br><small style='color:#555'>{parsed_auditors[0]['role']}</small>"
    aud2_html = f"{parsed_auditors[1]['display_name']}<br><small style='color:#555'>{parsed_auditors[1]['role']}</small>" if is_paired else "-"
    disp2 = curr_date if is_paired else "-"

    def _row(s1, s2):
        def cell(it, side):
            if not it: return "<td></td><td colspan='2'></td><td></td><td></td>"
            p = it["process"]; d_r = it["dur_real"]; d_e = it["dur_extra"]
            t = it["time"].replace(" - ","<br>")
            nc_style = "color:#c0392b;font-weight:700;" if "Previous audit" in p else ""
            lines = p.split("\n")
            p_html = "<br>".join(f"<span style='{nc_style}'>{ln}</span>" if ("Previous audit" in ln or "NC effective" in ln) else ln for ln in lines)
            dr_str = f"{d_r:.1f}" if d_r > 0 else "0"
            de_str = f"<span style='color:#c0392b;font-weight:700'>{d_e:.1f}</span>" if d_e > 0 else "0"
            return f"<td style='text-align:center;white-space:nowrap'>{t}</td><td colspan='2' style='text-align:left;padding:4px 8px'>{p_html}</td><td style='text-align:center'>{dr_str}</td><td style='text-align:center'>{de_str}</td>"
        return f"<tr style='height:32px'>{cell(s1,0)}{cell(s2,1)}</tr>"

    rows_html = ""
    for i in range(max_items):
        i1 = sched1[i] if i < len(sched1) else None
        i2 = sched2[i] if i < len(sched2) else None
        rows_html += _row(i1, i2)

    badge1 = "✅ 10시간 이내" if tot_dur1 <= 10 else f"❌ 10시간 초과"
    badge2 = "✅ 10시간 이내" if tot_dur2 <= 10 else f"❌ 10시간 초과"
    add2_str = f"{tot_add2:.1f}" if is_paired else "-"
    act2_str = f"{tot_act2:.1f}" if is_paired else "-"
    dur2_str = f"{tot_dur2:.1f}" if is_paired else "-"

    html = f"""<div style='font-family:system-ui,sans-serif;margin:20px 0;background:#fff;border:1px solid #d0d7de;border-radius:8px;padding:16px;box-shadow:0 2px 6px rgba(0,0,0,.05)'>
<div style='text-align:center;font-size:18px;font-weight:700;margin-bottom:4px'>AUDIT PLAN 심사계획서 (Day {day_idx})</div>
<div style='text-align:center;font-size:11px;color:#666;margin-bottom:12px'>To add up audit duration totals, click anywhere in the audit plan, press Ctrl+a, then select F9<br>심사기간 총계를 추가하려면 심사계획서의 아무 곳이나 클릭하고 Ctrl+a를 누른 다음 F9를 클릭하세요.</div>
<table style='width:100%;border-collapse:collapse;font-size:12px;border:1px solid #A6A6A6'>
<colgroup><col style='width:7%'><col style='width:5%'><col style='width:26%'><col style='width:5%'><col style='width:5%'><col style='width:7%'><col style='width:5%'><col style='width:26%'><col style='width:5%'><col style='width:5%'></colgroup>
<tr><th colspan='2' style='background:#FDE9D9;border:1px solid #A6A6A6;padding:4px 8px;text-align:left'>Date 심사일자</th><td colspan='3' style='border:1px solid #A6A6A6;text-align:center'>{curr_date}</td><th colspan='2' style='background:#FDE9D9;border:1px solid #A6A6A6;padding:4px 8px;text-align:left'>Date 심사일자</th><td colspan='3' style='border:1px solid #A6A6A6;text-align:center'>{disp2}</td></tr>
<tr><th colspan='2' style='background:#FDE9D9;border:1px solid #A6A6A6;padding:4px 8px;text-align:left'>Auditor Name/Role/<br>EA Code<br>심사원 이름/역할/<br>EA Code</th><td colspan='3' style='border:1px solid #A6A6A6;padding:4px 8px'>{aud1_html}</td><th colspan='2' style='background:#FDE9D9;border:1px solid #A6A6A6;padding:4px 8px;text-align:left'>Auditor Name/Role/<br>EA Code<br>심사원 이름/역할/<br>EA Code</th><td colspan='3' style='border:1px solid #A6A6A6;padding:4px 8px'>{aud2_html}</td></tr>
<tr><th colspan='2' style='background:#FDE9D9;border:1px solid #A6A6A6;padding:4px 8px;text-align:left'>Scope 인증범위</th><td colspan='8' style='border:1px solid #A6A6A6;padding:4px 8px'>{scope}</td></tr>
<tr style='background:#FDE9D9;font-weight:700;text-align:center'><th style='border:1px solid #A6A6A6'>TIME<br>시간</th><th colspan='2' style='border:1px solid #A6A6A6'>Location/Process<br>위치 / 프로세스</th><th colspan='2' style='border:1px solid #A6A6A6'>Duration<br>기간</th><th style='border:1px solid #A6A6A6'>TIME<br>시간</th><th colspan='2' style='border:1px solid #A6A6A6'>Location/Process<br>위치 / 프로세스</th><th colspan='2' style='border:1px solid #A6A6A6'>Duration<br>기간</th></tr>
{rows_html}
<tr style='font-weight:700'><td colspan='3' style='border:1px solid #A6A6A6;padding:4px 8px'>Total Daily Audit Time (hours)<br>일일 총 심사시간(hours)</td><td style='background:#FAE2D5;border:1px solid #A6A6A6;text-align:center'>{tot_act1:.1f}</td><td style='background:#FAE2D5;border:1px solid #A6A6A6;text-align:center'>{tot_add1:.1f}</td><td colspan='3' style='border:1px solid #A6A6A6;padding:4px 8px'>Total Daily Audit Time (hours)<br>일일 총 심사시간(hours)</td><td style='background:#FAE2D5;border:1px solid #A6A6A6;text-align:center'>{act2_str}</td><td style='background:#FAE2D5;border:1px solid #A6A6A6;text-align:center'>{add2_str}</td></tr>
<tr style='font-weight:700'><td colspan='5' style='border:1px solid #A6A6A6;padding:4px 8px'>Total Daily Additional Audit Time (hours)<br>일일 총 추가적 심사시간(hours)</td><td colspan='5' style='border:1px solid #A6A6A6;padding:4px 8px'>Total Daily Additional Audit Time (hours)<br>일일 총 추가적 심사시간(hours)</td></tr>
<tr style='font-weight:700'><td colspan='4' style='border:1px solid #A6A6A6;padding:4px 8px'>Total Daily Audit Duration (hours)<br>일일 총 심사기간(hours) <span style='font-size:11px;color:#0969da'>[{badge1}]</span></td><td style='background:#FAE2D5;border:1px solid #A6A6A6;text-align:center'>{tot_dur1:.1f}</td><td colspan='4' style='border:1px solid #A6A6A6;padding:4px 8px'>Total Daily Audit Duration (hours)<br>일일 총 심사기간(hours) <span style='font-size:11px;color:#0969da'>[{badge2 if is_paired else "-"}]</span></td><td style='background:#FAE2D5;border:1px solid #A6A6A6;text-align:center'>{dur2_str}</td></tr>
<tr style='background:#FDE9D9;font-size:11px'><td colspan='10' style='border:1px solid #A6A6A6;padding:6px 10px'>Audit trails will be developed based upon identified risk throughout the audit and as such timings and content may be subject to change<br>심사 진행은 심사 동안 파악된 리스크에 따라 변경될 것이며, 이에 따라 시기 및 내용이 변경될 수 있다.</td></tr>
</table></div>"""
    st.markdown(html, unsafe_allow_html=True)

# Word 다운로드
st.markdown("---")
st.header("📥 공식 Word 심사계획서 다운로드")
if DOCX_AVAILABLE:
    doc_buf = create_cara_word_document(client_name, scope,
        datetime.combine(audit_date_input, datetime.min.time()),
        parsed_auditors, all_schedules)
    safe_name = re.sub(r'[\\/*?:"<>|]', "", client_name).strip() or "CARA"
    st.download_button("📄 첨부해주신 공식 Word (.docx) 서식으로 다운로드",
        data=doc_buf, file_name=f"{safe_name}_IATF16949_심사계획서_{audit_date_input.strftime('%Y%m%d')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary", use_container_width=True)
    st.caption("💡 다운로드받은 Word 파일은 10열 표준 그리드 및 살구색(#FDE9D9) 헤더 서식과 100% 동일하게 레이아웃이 유지됩니다.")

st.divider()

# ══════════════════════════════════════════════
# 검증 리포트
# ══════════════════════════════════════════════
st.header("🔍 심사 계획 검증 데이터 및 시간·MD 통계 (내부 확인용)")
st.caption("💡 본 검증 자료는 심사계획서 생성 후 제조 비중 31%~34% 준수 여부 및 심사원별/전체 심사시간과 MD 배정을 점검하기 위한 화면 전용 분석 리포트입니다. (Word 출력물에는 포함되지 않습니다.)")

proc_df2 = gen["proc_df"]
mfg_names = proc_df2[proc_df2["제조프로세스여부"]==True]["프로세스명"].tolist()

total_base_h = sum(a["days_attending"] for a in parsed_auditors) * HOURS_PER_MD
total_extra_h = sum(nc.get("extra_hours",0) for nc in st.session_state.nc_list)

# 제조 심사시간
mfg_h = 0.0
for s in all_schedules:
    for _, r in s.iterrows():
        if any(m in r["process"] for m in mfg_names):
            mfg_h += r["dur_real"]
mfg_pct = mfg_h / total_base_h * 100 if total_base_h > 0 else 0

r1c1, r1c2, r1c3, r1c4 = st.columns(4)
with r1c1:
    st.metric("🏭 제조 심사시간 및 비율", f"{mfg_h:.1f} 시간 ({mfg_pct:.1f}%)", "IATF 16949 기준: 31%~34%")
with r1c2:
    st.metric("⏱️ 전체 기본 심사시간", f"{total_base_h:.1f} 시간", f"참여 {sum(a['days_attending'] for a in parsed_auditors):.1f} MD (1일 8h 기준)")
with r1c3:
    st.metric("➕ 추가 심사시간 (시정조치)", f"{total_extra_h:.1f} 시간", "부적합 검증 추가 배정")
with r1c4:
    total_all = total_base_h + total_extra_h
    total_md_all = total_all / HOURS_PER_MD
    st.metric("📊 전체 총 심사시간 / 환산 MD", f"{total_all:.1f} 시간", f"총 {total_md_all:.2f} MD 환산")

st.subheader("1️⃣ 제조 프로세스 심사 시간 및 전체 심사 시간과의 비율 검증")
if MIN_MFG_RATIO*100 <= mfg_pct <= MAX_MFG_RATIO*100:
    st.success(f"🟢 **적합 (31% ~ 34% 최적 준수)** — 전체 기본 심사시간({total_base_h:.1f}h) 중 제조 심사시간이 **{mfg_h:.1f}시간({mfg_pct:.1f}%)** 배정되어, IATF 16949 핵심 규정(31%~34%)을 완벽히 충족합니다.")
else:
    st.warning(f"🟡 **주의** — 제조 비중 {mfg_pct:.1f}% (목표: 31~34%). 조정이 필요합니다.")

# 심사원별 집계
st.subheader("2️⃣ & 3️⃣ 각 심사원별 및 전체 심사원 심사시간·MD 집계표")
rows_tbl = []
for i, (aud, sc) in enumerate(zip(parsed_auditors, all_schedules)):
    base = sc["dur_real"].sum()
    extra = sc["dur_extra"].sum()
    mfg_i = sum(r["dur_real"] for _, r in sc.iterrows() if any(m in r["process"] for m in mfg_names))
    rows_tbl.append({
        "구분": "선임심사원 (Lead)" if i==0 else f"심사원 #{i+1}",
        "심사원명": aud["display_name"],
        "배정 역할 및 EAC": aud["role"],
        "참여일수 (MD)": f"{aud['days_attending']:.1f} MD",
        "기본 심사시간": f"{base:.1f} 시간",
        "추가 심사시간": f"{extra:.1f}h" if extra else "-",
        "총 심사시간": f"{base+extra:.1f} 시간",
        "총 심사 MD": f"{(base+extra)/HOURS_PER_MD:.2f} MD",
        "제조 심사시간": f"{mfg_i:.1f}h ({mfg_i/base*100:.1f}%)" if base and mfg_i else "-",
    })
rows_tbl.append({
    "구분":"전체 합계","심사원명":f"총 {len(parsed_auditors)}명","배정 역할 및 EAC":"-",
    "참여일수 (MD)":f"{sum(a['days_attending'] for a in parsed_auditors):.1f} MD",
    "기본 심사시간":f"{total_base_h:.1f} 시간",
    "추가 심사시간":f"{total_extra_h:.1f}h" if total_extra_h else "-",
    "총 심사시간":f"{total_all:.1f} 시간",
    "총 심사 MD":f"{total_md_all:.2f} MD",
    "제조 심사시간":f"{mfg_h:.1f}h ({mfg_pct:.1f}%)",
})
st.dataframe(pd.DataFrame(rows_tbl), use_container_width=True, hide_index=True)
st.caption("※ 1 MD는 8.0 심사 시간 기준이며, 추가 심사시간(이전 부적합 Minor NC 검증 등)은 1일 10시간 한도 내에서 가산되어 총 심사시간 및 총 심사 MD에 합산됩니다.")

with st.sidebar:
    st.markdown("## ⚙️ 심사 운영 규칙")
    st.markdown(f"- **1일 기본**: {HOURS_PER_MD:.0f}h (1 MD)")
    st.markdown(f"- **1일 최대**: {MAX_DAILY_HOURS:.0f}h")
    st.markdown(f"- **제조 비중**: {MIN_MFG_RATIO*100:.0f}~{MAX_MFG_RATIO*100:.0f}%")
