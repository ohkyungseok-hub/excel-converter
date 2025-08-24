# app_upload_fix.py
# 실행: streamlit run app_upload_fix.py
# 필요: pip install streamlit pandas openpyxl

import io
import re
import json
import pandas as pd
import streamlit as st

st.set_page_config(page_title="엑셀 양식 변환기 (1→2)", layout="centered")

st.title("엑셀 양식 변환기 (1 → 2)")
st.caption("라오라 / 쿠팡 / 스마트스토어(고정) / 떠리몰(고정) 형식을 2번 템플릿으로 변환합니다. (전화번호 0 보존)")

# -------------------------- Helpers --------------------------
def excel_col_to_index(col_letters: str) -> int:
    col_letters = str(col_letters).strip().upper()
    if not re.fullmatch(r"[A-Z]+", col_letters):
        raise ValueError(f"Invalid Excel column letters: {col_letters}")
    idx = 0
    for ch in col_letters:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1  # 0-based

def index_to_excel_col(n: int) -> str:
    s = ""
    n += 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(r + 65) + s
    return s

def excel_letters(max_cols=104):
    return [index_to_excel_col(i) for i in range(max_cols)]

def read_first_sheet_template(file) -> pd.DataFrame:
    """템플릿(2.xlsx)은 일반적으로 읽기"""
    return pd.read_excel(file, sheet_name=0, header=0, engine="openpyxl")

def read_first_sheet_source_as_text(file) -> pd.DataFrame:
    """소스는 전 컬럼을 문자열로 읽어 전화번호 앞 0 보존"""
    return pd.read_excel(
        file,
        sheet_name=0,
        header=0,
        engine="openpyxl",
        dtype=str,
        keep_default_na=False,  # 빈값을 NaN 대신 빈 문자열로 유지
    )

def ensure_mapping_initialized(template_columns, default_mapping):
    m = st.session_state.get("mapping")
    if not isinstance(m, dict):
        m = {}
    synced = {k: str(v).upper() for k, v in m.items() if k in template_columns and v}
    for k in template_columns:
        if k not in synced and k in default_mapping:
            synced[k] = default_mapping[k]
    st.session_state["mapping"] = synced
    return st.session_state["mapping"]

# -------------------- Defaults --------------------
DEFAULT_TEMPLATE_COLUMNS = [
    "주문번호",
    "받는분 이름",
    "받는분 주소",
    "받는분 전화번호",
    "상품명",
    "수량",
    "메모",
]

# 라오라 기본 매핑 (열 문자)
DEFAULT_MAPPING = {
    "주문번호": "D",
    "받는분 이름": "AM",
    "받는분 주소": "AP",
    "받는분 전화번호": "AN",
    "상품명": "S",
    "수량": "V",
    "메모": "AR",
}

# 쿠팡 고정 매핑 (열 문자) — 주문번호 C
COUPANG_MAPPING = {
    "주문번호": "C",
    "받는분 이름": "AA",
    "받는분 주소": "AD",
    "받는분 전화번호": "AB",
    "상품명": "P",
    "수량": "W",
    "메모": "AE",
}

# 스마트스토어 고정 매핑 (열 문자)
# B: 주문번호, L: 수취인명, AP: 통합배송지, AN: 수취인연락처, Q & S: 상품명(연결), U: 수량, AU: 배송메시지
SMARTSTORE_FIXED_LETTER_MAPPING = {
    "주문번호": "B",
    "받는분 이름": "L",
    "받는분 주소": "AP",
    "받는분 전화번호": "AN",
    "상품명_Q": "Q",
    "상품명_S": "S",
    "수량": "U",
    "메모": "AU",
}

# 떠리몰 고정 매핑 (열 문자)
# H: 주문번호, AB: 수령자명, AE: 주소, AC: 수령자연락처, V: 옵션명:옵션값(→상품명), Y: 수량, AA: 배송메시지
TTARIMALL_FIXED_LETTER_MAPPING = {
    "주문번호": "H",
    "받는분 이름": "AB",
    "받는분 주소": "AE",
    "받는분 전화번호": "AC",
    "상품명": "V",
    "수량": "Y",
    "메모": "AA",
}

# -------------------------- Sidebar --------------------------
st.sidebar.header("템플릿 옵션")
use_uploaded_template = st.sidebar.checkbox("템플릿(2.xlsx) 직접 업로드", value=False)
max_letter_cols = st.sidebar.slider(
    "라오라용 최대 열 범위(Excel 문자)",
    min_value=52,
    max_value=156,
    value=104,
    step=26,
    help="라오라 매핑 드롭다운의 열 문자 개수",
)
st.sidebar.divider()
st.sidebar.subheader("라오라 매핑 저장/불러오기")
mapping_upload = st.sidebar.file_uploader("매핑 JSON 불러오기 (라오라)", type=["json"], key="mapping_json")
prepare_download = st.sidebar.button("현재 라오라 매핑 JSON 다운로드 준비")

# -------------------------- 템플릿 설정 (공용) --------------------------
st.subheader("템플릿 설정 (2.xlsx)")
tpl_df = None
if use_uploaded_template:
    tpl_file = st.file_uploader("2와 같은 템플릿 파일 업로드 (예: 2.xlsx)", type=["xlsx"], key="tpl")
    if tpl_file:
        try:
            tpl_df = read_first_sheet_template(tpl_file)
            st.success(f"템플릿 업로드 완료. 컬럼 수: {len(tpl_df.columns)}")
        except Exception as e:
            st.warning(f"템플릿 파일을 읽는 중 오류가 발생했습니다: {e}")
            tpl_df = None
else:
    tpl_df = pd.DataFrame(columns=DEFAULT_TEMPLATE_COLUMNS)
    st.info("업로드된 템플릿이 없으므로 기본 템플릿을 사용합니다. (주문번호, 받는분 이름, 받는분 주소, 받는분 전화번호, 상품명, 수량, 메모)")

template_columns = list(tpl_df.columns) if tpl_df is not None else []

# ======================================================================
# 1) 라오라 파일 변환 (열 문자 매핑)
# ======================================================================
st.markdown("## 라오라 파일 변환")

current_mapping = ensure_mapping_initialized(template_columns, DEFAULT_MAPPING)
letters = excel_letters(max_letter_cols)

if mapping_upload is not None:
    try:
        loaded = json.load(mapping_upload)
        if not isinstance(loaded, dict):
            raise ValueError("JSON 루트가 객체(dict)가 아닙니다.")
        new_map = {}
        for k, v in loaded.items():
            if k in template_columns and isinstance(v, str) and re.fullmatch(r"[A-Za-z]+", v):
                new_map[k] = v.upper()
        for k in template_columns:
            if k not in new_map:
                new_map[k] = current_mapping.get(k, DEFAULT_MAPPING.get(k, ""))
        st.session_state["mapping"] = new_map
        current_mapping = new_map
        st.success("라오라 매핑 JSON을 불러왔습니다.")
    except Exception as e:
        st.warning(f"라오라 매핑 JSON 불러오기 실패: {e}")

edited_mapping = {}
with st.form("mapping_form_laora"):
    for col in template_columns:
        default_val = current_mapping.get(col, "")
        if default_val not in letters:
            default_val = ""
        options = [""] + letters
        sel = st.selectbox(
            f"{col} ⟶ 1.xlsx(라오라) 열 문자 선택",
            options=options,
            index=(options.index(default_val) if default_val in options else 0),
            key=f"map_laora_{col}",
        )
        edited_mapping[col] = sel
    if st.form_submit_button("라오라 매핑 저장"):
        st.session_state["mapping"] = {k: v for k, v in edited_mapping.items() if v}
        current_mapping = st.session_state["mapping"]
        st.success("라오라 매핑을 저장했습니다.")

if prepare_download:
    mapping_bytes = json.dumps(current_mapping, ensure_ascii=False, indent=2).encode("utf-8")
    st.download_button(
        label="현재 라오라 매핑 JSON 다운로드",
        data=mapping_bytes,
        file_name="mapping_laora.json",
        mime="application/json",
    )

st.subheader("라오라 소스 파일 업로드")
src_file_laora = st.file_uploader("라오라 형식의 파일 업로드 (예: 1.xlsx)", type=["xlsx"], key="src_laora")
run_laora = st.button("라오라 변환 실행")
if run_laora:
    if not src_file_laora:
        st.error("라오라 소스 파일을 업로드해 주세요.")
    elif tpl_df is None or len(template_columns) == 0:
        st.error("유효한 템플릿이 필요합니다.")
    else:
        try:
            df_src = read_first_sheet_source_as_text(src_file_laora)
        except Exception as e:
            st.exception(RuntimeError(f"라오라 소스 파일을 읽는 중 오류: {e}"))
        else:
            result = pd.DataFrame(index=range(len(df_src)), columns=template_columns)
            mapping = st.session_state.get("mapping", {})
            if not isinstance(mapping, dict) or not mapping:
                st.error("라오라 매핑이 없습니다. 먼저 저장해 주세요.")
            else:
                src_cols_by_index = list(df_src.columns)
                resolved_map = {}
                try:
                    for tpl_header, xl_letters in mapping.items():
                        if not xl_letters:
                            continue
                        idx = excel_col_to_index(xl_letters)
                        if idx >= len(src_cols_by_index):
                            raise IndexError(
                                f"소스 파일에 {xl_letters} 열(0-based index {idx})이 존재하지 않습니다. "
                                f"소스 컬럼 수: {len(src_cols_by_index)}"
                            )
                        resolved_map[tpl_header] = src_cols_by_index[idx]
                except Exception as e:
                    st.exception(RuntimeError(f"라오라 매핑 인덱스 계산 중 오류: {e}"))
                else:
                    for tpl_header, src_colname in resolved_map.items():
                        try:
                            if tpl_header == "수량":
                                result[tpl_header] = pd.to_numeric(df_src[src_colname], errors="coerce")
                            elif tpl_header == "받는분 전화번호":
                                series = df_src[src_colname].astype(str)
                                result[tpl_header] = series.where(series.str.lower() != "nan", "")
                            else:
                                result[tpl_header] = df_src[src_colname]
                        except KeyError:
                            st.warning(f"소스 컬럼 '{src_colname}'(매핑: {tpl_header})을(를) 찾을 수 없습니다. 해당 필드는 비워집니다.")

                    # 템플릿 숫자형 정렬(전화번호 제외)
                    for col in template_columns:
                        if col in tpl_df.columns and tpl_df[col].notna().any():
                            if pd.api.types.is_numeric_dtype(tpl_df[col]) and col != "받는분 전화번호":
                                result[col] = pd.to_numeric(result[col], errors="coerce")

                    st.success(f"라오라 변환 완료: 총 {len(result)}행")
                    st.dataframe(result.head(50))

                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                        out_df = result[template_columns + [c for c in result.columns if c not in template_columns]]
                        out_df.to_excel(writer, index=False)
                    st.download_button(
                        label="라오라 변환 결과 다운로드 (output_laora.xlsx)",
                        data=buffer.getvalue(),
                        file_name="output_laora.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

st.markdown("---")

# ======================================================================
# 2) 쿠팡 파일 변환 (고정 매핑)
# ======================================================================
st.markdown("## 쿠팡 파일 변환")

with st.expander("쿠팡 → 템플릿 매핑 보기", expanded=False):
    st.markdown(
        """
        **쿠팡 소스열 → 템플릿 컬럼**  
        - `C` → **주문번호**  
        - `AA` → **받는분 이름**  
        - `AD` → **받는분 주소**  
        - `AB` → **받는분 전화번호**  
        - `P` → **상품명** (최초등록상품명/옵션명)  
        - `W` → **수량** (구매수)  
        - `AE` → **메모** (배송메시지)
        """
    )

st.subheader("쿠팡 소스 파일 업로드")
src_file_coupang = st.file_uploader("쿠팡 형식의 파일 업로드 (예: 쿠팡.xlsx)", type=["xlsx"], key="src_coupang")
run_coupang = st.button("쿠팡 변환 실행")
if run_coupang:
    if not src_file_coupang:
        st.error("쿠팡 소스 파일을 업로드해 주세요.")
    elif tpl_df is None or len(template_columns) == 0:
        st.error("유효한 템플릿이 필요합니다.")
    else:
        try:
            df_src_cp = read_first_sheet_source_as_text(src_file_coupang)
        except Exception as e:
            st.exception(RuntimeError(f"쿠팡 소스 파일을 읽는 중 오류: {e}"))
        else:
            result_cp = pd.DataFrame(index=range(len(df_src_cp)), columns=template_columns)
            mapping_cp = COUPANG_MAPPING.copy()

            src_cols_by_index_cp = list(df_src_cp.columns)
            resolved_map_cp = {}
            try:
                for tpl_header, xl_letters in mapping_cp.items():
                    idx = excel_col_to_index(xl_letters)
                    if idx >= len(src_cols_by_index_cp):
                        raise IndexError(
                            f"쿠팡 소스에 {xl_letters} 열(0-based index {idx})이 존재하지 않습니다. "
                            f"소스 컬럼 수: {len(src_cols_by_index_cp)}"
                        )
                    resolved_map_cp[tpl_header] = src_cols_by_index_cp[idx]
            except Exception as e:
                st.exception(RuntimeError(f"쿠팡 매핑 인덱스 계산 중 오류: {e}"))
            else:
                for tpl_header, src_colname in resolved_map_cp.items():
                    try:
                        if tpl_header == "수량":
                            result_cp[tpl_header] = pd.to_numeric(df_src_cp[src_colname], errors="coerce")
                        elif tpl_header == "받는분 전화번호":
                            series = df_src_cp[src_colname].astype(str)
                            result_cp[tpl_header] = series.where(series.str.lower() != "nan", "")
                        else:
                            result_cp[tpl_header] = df_src_cp[src_colname]
                    except KeyError:
                        st.warning(f"[쿠팡] 소스 컬럼 '{src_colname}'(매핑: {tpl_header})을(를) 찾을 수 없습니다. 해당 필드는 비워집니다.")

                # 템플릿 숫자형 정렬(전화번호 제외)
                for col in template_columns:
                    if col in tpl_df.columns and tpl_df[col].notna().any():
                        if pd.api.types.is_numeric_dtype(tpl_df[col]) and col != "받는분 전화번호":
                            result_cp[col] = pd.to_numeric(result_cp[col], errors="coerce")

                st.success(f"쿠팡 변환 완료: 총 {len(result_cp)}행")
                st.dataframe(result_cp.head(50))

                buffer_cp = io.BytesIO()
                with pd.ExcelWriter(buffer_cp, engine="openpyxl") as writer:
                    out_df_cp = result_cp[template_columns + [c for c in result_cp.columns if c not in template_columns]]
                    out_df_cp.to_excel(writer, index=False)
                st.download_button(
                    label="쿠팡 변환 결과 다운로드 (output_coupang.xlsx)",
                    data=buffer_cp.getvalue(),
                    file_name="output_coupang.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

st.markdown("---")

# ======================================================================
# 3) 스마트스토어 파일 변환 (고정 매핑: 열 문자)
# ======================================================================
st.markdown("## 스마트스토어 파일 변환 (고정 매핑: 열 문자)")

with st.expander("스마트스토어(고정) → 템플릿 매핑 보기", expanded=False):
    st.markdown(
        """
        **스마트스토어 소스열 → 템플릿 컬럼**  
        - `B` → **주문번호**  
        - `L` → **받는분 이름** (수취인명)  
        - `AP` → **받는분 주소** (통합배송지)  
        - `AN` → **받는분 전화번호** (수취인연락처)  
        - `Q & S` → **상품명** (두 열을 그대로 연결)  
        - `U` → **수량**  
        - `AU` → **메모** (배송메시지)
        """
    )

st.subheader("스마트스토어 소스 파일 업로드 (고정 매핑)")
src_file_ss_fixed = st.file_uploader("스마트스토어 형식의 파일 업로드 (예: 스마트스토어.xlsx)", type=["xlsx"], key="src_smartstore_fixed")

run_ss_fixed = st.button("스마트스토어 변환 실행 (고정 매핑)")
if run_ss_fixed:
    if not src_file_ss_fixed:
        st.error("스마트스토어 소스 파일을 업로드해 주세요.")
    elif tpl_df is None or len(template_columns) == 0:
        st.error("유효한 템플릿이 필요합니다.")
    else:
        try:
            df_ss_fixed = read_first_sheet_source_as_text(src_file_ss_fixed)
        except Exception as e:
            st.exception(RuntimeError(f"스마트스토어 소스 파일을 읽는 중 오류: {e}"))
        else:
            result_ss_fixed = pd.DataFrame(index=range(len(df_ss_fixed)), columns=template_columns)

            src_cols_by_index_ss = list(df_ss_fixed.columns)

            def resolve(letter: str) -> str:
                idx = excel_col_to_index(letter)
                if idx >= len(src_cols_by_index_ss):
                    raise IndexError(
                        f"스마트스토어 소스에 {letter} 열(0-based index {idx})이 없습니다. "
                        f"소스 컬럼 수: {len(src_cols_by_index_ss)}"
                    )
                return src_cols_by_index_ss[idx]

            try:
                col_order = resolve(SMARTSTORE_FIXED_LETTER_MAPPING["주문번호"])
                col_name  = resolve(SMARTSTORE_FIXED_LETTER_MAPPING["받는분 이름"])
                col_addr  = resolve(SMARTSTORE_FIXED_LETTER_MAPPING["받는분 주소"])
                col_phone = resolve(SMARTSTORE_FIXED_LETTER_MAPPING["받는분 전화번호"])
                col_q     = resolve(SMARTSTORE_FIXED_LETTER_MAPPING["상품명_Q"])
                col_s     = resolve(SMARTSTORE_FIXED_LETTER_MAPPING["상품명_S"])
                col_qty   = resolve(SMARTSTORE_FIXED_LETTER_MAPPING["수량"])
                col_memo  = resolve(SMARTSTORE_FIXED_LETTER_MAPPING["메모"])
            except Exception as e:
                st.exception(RuntimeError(f"스마트스토어 고정 매핑 인덱스 계산 중 오류: {e}"))
            else:
                result_ss_fixed["주문번호"] = df_ss_fixed[col_order]
                result_ss_fixed["받는분 이름"] = df_ss_fixed[col_name]
                result_ss_fixed["받는분 주소"] = df_ss_fixed[col_addr]
                series_phone = df_ss_fixed[col_phone].astype(str)
                result_ss_fixed["받는분 전화번호"] = series_phone.where(series_phone.str.lower() != "nan", "")
                q_series = df_ss_fixed[col_q].astype(str).where(lambda s: s.str.lower() != "nan", "")
                s_series = df_ss_fixed[col_s].astype(str).where(lambda s: s.str.lower() != "nan", "")
                result_ss_fixed["상품명"] = q_series.fillna("") + s_series.fillna("")
                result_ss_fixed["수량"] = pd.to_numeric(df_ss_fixed[col_qty], errors="coerce")
                result_ss_fixed["메모"] = df_ss_fixed[col_memo]

                for col in template_columns:
                    if col in tpl_df.columns and tpl_df[col].notna().any():
                        if pd.api.types.is_numeric_dtype(tpl_df[col]) and col != "받는분 전화번호":
                            result_ss_fixed[col] = pd.to_numeric(result_ss_fixed[col], errors="coerce")

                st.success(f"스마트스토어(고정) 변환 완료: 총 {len(result_ss_fixed)}행")
                st.dataframe(result_ss_fixed.head(50))

                buffer_ss_fixed = io.BytesIO()
                with pd.ExcelWriter(buffer_ss_fixed, engine="openpyxl") as writer:
                    out_df_ss_fixed = result_ss_fixed[template_columns + [c for c in result_ss_fixed.columns if c not in template_columns]]
                    out_df_ss_fixed.to_excel(writer, index=False)
                st.download_button(
                    label="스마트스토어(고정) 변환 결과 다운로드 (output_smartstore_fixed.xlsx)",
                    data=buffer_ss_fixed.getvalue(),
                    file_name="output_smartstore_fixed.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

st.markdown("---")

# ======================================================================
# 4) 떠리몰 파일 변환 (고정 매핑: 열 문자)
# ======================================================================
st.markdown("## 떠리몰 파일 변환 (고정 매핑: 열 문자)")

with st.expander("떠리몰(고정) → 템플릿 매핑 보기", expanded=False):
    st.markdown(
        """
        **떠리몰 소스열 → 템플릿 컬럼**  
        - `H` → **주문번호**  
        - `AB` → **받는분 이름** (수령자명)  
        - `AE` → **받는분 주소** (주소)  
        - `AC` → **받는분 전화번호** (수령자연락처)  
        - `V` → **상품명** (옵션명:옵션값)  
        - `Y` → **수량**  
        - `AA` → **메모** (배송메시지)
        """
    )

st.subheader("떠리몰 소스 파일 업로드 (고정 매핑)")
src_file_ttarimall = st.file_uploader("떠리몰 형식의 파일 업로드 (예: 떠리몰.xlsx)", type=["xlsx"], key="src_ttarimall")

run_ttarimall = st.button("떠리몰 변환 실행 (고정 매핑)")
if run_ttarimall:
    if not src_file_ttarimall:
        st.error("떠리몰 소스 파일을 업로드해 주세요.")
    elif tpl_df is None or len(template_columns) == 0:
        st.error("유효한 템플릿이 필요합니다.")
    else:
        try:
            df_tm = read_first_sheet_source_as_text(src_file_ttarimall)
        except Exception as e:
            st.exception(RuntimeError(f"떠리몰 소스 파일을 읽는 중 오류: {e}"))
        else:
            result_tm = pd.DataFrame(index=range(len(df_tm)), columns=template_columns)

            src_cols_by_index_tm = list(df_tm.columns)

            def resolve(letter: str) -> str:
                idx = excel_col_to_index(letter)
                if idx >= len(src_cols_by_index_tm):
                    raise IndexError(
                        f"떠리몰 소스에 {letter} 열(0-based index {idx})이 없습니다. "
                        f"소스 컬럼 수: {len(src_cols_by_index_tm)}"
                    )
                return src_cols_by_index_tm[idx]

            try:
                col_order = resolve(TTARIMALL_FIXED_LETTER_MAPPING["주문번호"])
                col_name  = resolve(TTARIMALL_FIXED_LETTER_MAPPING["받는분 이름"])
                col_addr  = resolve(TTARIMALL_FIXED_LETTER_MAPPING["받는분 주소"])
                col_phone = resolve(TTARIMALL_FIXED_LETTER_MAPPING["받는분 전화번호"])
                col_prod  = resolve(TTARIMALL_FIXED_LETTER_MAPPING["상품명"])
                col_qty   = resolve(TTARIMALL_FIXED_LETTER_MAPPING["수량"])
                col_memo  = resolve(TTARIMALL_FIXED_LETTER_MAPPING["메모"])
            except Exception as e:
                st.exception(RuntimeError(f"떠리몰 고정 매핑 인덱스 계산 중 오류: {e}"))
            else:
                # 채우기
                result_tm["주문번호"] = df_tm[col_order]
                result_tm["받는분 이름"] = df_tm[col_name]
                result_tm["받는분 주소"] = df_tm[col_addr]
                # 전화번호: 문자열(0 보존)
                series_phone = df_tm[col_phone].astype(str)
                result_tm["받는분 전화번호"] = series_phone.where(series_phone.str.lower() != "nan", "")
                # 상품명: V열(옵션명:옵션값) 그대로 사용
                prod_series = df_tm[col_prod].astype(str).where(lambda s: s.str.lower() != "nan", "")
                result_tm["상품명"] = prod_series
                # 수량
                result_tm["수량"] = pd.to_numeric(df_tm[col_qty], errors="coerce")
                # 메모
                result_tm["메모"] = df_tm[col_memo]

                # 템플릿 숫자형 정렬(전화번호 제외)
                for col in template_columns:
                    if col in tpl_df.columns and tpl_df[col].notna().any():
                        if pd.api.types.is_numeric_dtype(tpl_df[col]) and col != "받는분 전화번호":
                            result_tm[col] = pd.to_numeric(result_tm[col], errors="coerce")

                st.success(f"떠리몰(고정) 변환 완료: 총 {len(result_tm)}행")
                st.dataframe(result_tm.head(50))

                buffer_tm = io.BytesIO()
                with pd.ExcelWriter(buffer_tm, engine="openpyxl") as writer:
                    out_df_tm = result_tm[template_columns + [c for c in result_tm.columns if c not in template_columns]]
                    out_df_tm.to_excel(writer, index=False)
                st.download_button(
                    label="떠리몰(고정) 변환 결과 다운로드 (output_ttarimall.xlsx)",
                    data=buffer_tm.getvalue(),
                    file_name="output_ttarimall.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

st.markdown("---")
st.caption("라오라 / 쿠팡 / 스마트스토어(고정) / 떠리몰(고정) 외 양식도 추가 가능합니다. 규칙만 알려주시면 바로 넣어드릴게요.")
