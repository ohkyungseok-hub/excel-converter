# app_upload_fix.py
# 실행: streamlit run app_upload_fix.py
# 필요: pip install streamlit pandas openpyxl
# (.xls 읽기 필요 시) pip install "xlrd==1.2.0"

import io
import re
import json
import zipfile
from datetime import datetime
from typing import Optional, List

import pandas as pd
import streamlit as st

st.set_page_config(page_title="황지후의 발주 대작전 (1→2)", layout="centered")

st.title("황지후의 발주 대작전 (1 → 2)")
st.caption("라오라 / 쿠팡 / 스마트스토어(키워드) / 떠리몰(키워드 S&V 규칙) 형식을 2번 템플릿으로 변환합니다. (전화번호 0 보존)")

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
        keep_default_na=False,
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

def norm_header(s: str) -> str:
    return re.sub(r"[\s\(\)\[\]{}:：/\\\-]", "", str(s).strip().lower())

# ★ CSV에서 Excel이 숫자로 오인하지 않도록 텍스트 보호
def _guard_excel_text(s: str) -> str:
    """
    Excel이 CSV를 열 때 숫자로 오인하지 않도록 '="값"' 형태로 감싸기.
    이미 ="..." 형태면 중복 적용하지 않음.
    """
    s = "" if s is None else str(s)
    if s == "" or s.startswith('="'):
        return s
    return f'="{s}"'

# -------------------- CSV 출력 설정(구분자/인코딩) --------------------
CSV_SEPARATORS = {
    "쉼표(,)": ",",
    "세미콜론(;)": ";",
    "탭(\\t)": "\t",
    "파이프(|)": "|",
}
CSV_ENCODINGS = {
    "UTF-8-SIG (권장)": "utf-8-sig",
    "UTF-8 (BOM 없음)": "utf-8",
    "CP949 (윈도우)": "cp949",
    "EUC-KR": "euc-kr",
}

def _get_csv_prefs():
    sep = st.session_state.get("csv_sep", ",")
    enc = st.session_state.get("csv_encoding", "utf-8-sig")
    label_sep = st.session_state.get("csv_sep_label", "쉼표(,)")
    label_enc = st.session_state.get("csv_enc_label", "UTF-8-SIG (권장)")
    return sep, enc, label_sep, label_enc

def download_df(
    df: pd.DataFrame,
    base_label: str,
    filename_stem: str,
    widget_key: str,
    sheet_name: Optional[str] = None,
    csv_sep_override: Optional[str] = None,      # ★ 추가
    csv_encoding_override: Optional[str] = None, # ★ 추가
):
    """CSV 버튼을 먼저, 그 다음에 XLSX 버튼을 보여주는 다운로드 위젯."""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    col_csv, col_xlsx = st.columns(2)

    # 현재 CSV 설정 불러오기 (오버라이드 우선)
    def _labels_from_sep(sep: str) -> str:
        return {",": "쉼표(,)", ";": "세미콜론(;)", "\t": "탭(\\t)", "|": "파이프(|)"}\
               .get(sep, f"사용자({repr(sep)})")
    def _labels_from_enc(enc: str) -> str:
        rev = {v: k for k, v in CSV_ENCODINGS.items()}
        return rev.get(enc, enc)

    default_sep, default_enc, default_sep_label, default_enc_label = _get_csv_prefs()
    csv_sep = csv_sep_override if csv_sep_override is not None else default_sep
    csv_enc = csv_encoding_override if csv_encoding_override is not None else default_enc
    label_sep = _labels_from_sep(csv_sep)
    label_enc = _labels_from_enc(csv_enc)

    # CSV 버튼 (전화번호 보호: ="010...")
    with col_csv:
        df_safe = df.copy()
        phone_like_cols = [c for c in df_safe.columns if re.search(r"(전화번호|연락처|휴대폰)", str(c))]
        for c in phone_like_cols:
            df_safe[c] = df_safe[c].astype(str).map(_guard_excel_text)

        csv_str = df_safe.to_csv(index=False, sep=csv_sep, lineterminator="\n")
        csv_bytes = csv_str.encode(csv_enc, errors="replace")
        st.download_button(
            label=f"{base_label} (CSV · {label_sep} · {label_enc})",
            data=csv_bytes,
            file_name=f"{filename_stem}_{ts}.csv",
            mime="text/csv",
            key=f"btn_{widget_key}_csv",
            help="선택한(또는 강제된) 구분자/인코딩으로 CSV 저장합니다.",
        )

    # XLSX 버튼
    with col_xlsx:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            if sheet_name:
                df.to_excel(writer, index=False, sheet_name=sheet_name)
            else:
                df.to_excel(writer, index=False)
        st.download_button(
            label=f"{base_label} (XLSX)",
            data=buf.getvalue(),
            file_name=f"{filename_stem}_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"btn_{widget_key}_xlsx",
            help="서식 유지가 필요한 경우 XLSX로 저장하세요.",
        )

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

# 라오 기본 매핑 (열 문자)
DEFAULT_MAPPING = {
    "주문번호": "A",
    "받는분 이름": "I",
    "받는분 주소": "L",
    "받는분 전화번호": "J",
    "상품명": "D",
    "수량": "G",
    "메모": "M",
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

# 스마트스토어 키워드 매핑용 후보
SS_NAME_MAP = {
    "주문번호": ["주문번호"],
    "받는분 이름": ["수취인명"],
    "받는분 주소": ["통합배송지"],
    "받는분 전화번호": ["수취인연락처1", "수취인연락처", "수취인휴대폰", "연락처1"],
    "상품명_left": ["상품명"],
    "상품명_right": ["옵션정보", "옵션명", "옵션내용"],
    "수량": ["수량", "구매수량"],
    "메모": ["배송메세지", "배송메시지", "배송요청사항"],
}

# ★ 떠리몰 키워드 매핑용 후보 (S & V 규칙)
TTARIMALL_NAME_MAP = {
    "주문번호": ["주문번호", "주문ID", "주문코드", "주문번호1"],
    "받는분 이름": ["수령자명", "받는분", "수취인명", "수령자"],
    "받는분 주소": ["주소", "수령자주소", "배송지주소", "통합배송지"],
    "받는분 전화번호": ["수령자연락처", "연락처", "휴대폰", "전화번호", "연락처1"],
    "상품명_S": ["상품명(S)", "상품명_S", "상품명S", "판매상품명", "상품명"],
    "상품명_V": ["옵션명:옵션값", "옵션", "옵션명", "옵션정보", "옵션내용", "옵션값", "상품옵션"],
    "수량": ["수량", "구매수", "주문수량"],
    "메모": ["배송메시지", "배송메세지", "배송요청사항", "메모"],
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

# ★ CSV 출력 설정
st.sidebar.divider()
st.sidebar.header("CSV 출력 설정")
sep_label = st.sidebar.selectbox("구분자", list(CSV_SEPARATORS.keys()), index=0, help="엑셀에서 쉼표로 인식하게 하려면 '쉼표(,)'를 선택하세요.")
enc_label = st.sidebar.selectbox("인코딩", list(CSV_ENCODINGS.keys()), index=0, help="엑셀 호환에는 보통 'UTF-8-SIG (권장)'을 사용합니다.")
st.session_state["csv_sep_label"] = sep_label
st.session_state["csv_enc_label"] = enc_label
st.session_state["csv_sep"] = CSV_SEPARATORS[sep_label]
st.session_state["csv_encoding"] = CSV_ENCODINGS[enc_label]

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

                    out_df = result[template_columns + [c for c in result.columns if c not in template_columns]]
                    download_df(out_df, "라오라 변환 결과 다운로드", "라오 3pl발주용", "laora_conv")

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

                out_df_cp = result_cp[template_columns + [c for c in result_cp.columns if c not in template_columns]]
                download_df(out_df_cp, "쿠팡 변환 결과 다운로드", "쿠팡 3pl발주용", "coupang_conv")

st.markdown("---")

# ======================================================================
# 3) 스마트스토어 파일 변환 (키워드 매핑)
# ======================================================================
st.markdown("## 스마트스토어 파일 변환 (키워드 매핑)")

with st.expander("스마트스토어(키워드) → 템플릿 매핑 보기", expanded=False):
    st.markdown(
        """
        **스마트스토어 컬럼명(헤더) → 템플릿 컬럼**  
        - `주문번호` → **주문번호**  
        - `수취인명` → **받는분 이름**  
        - `통합배송지` → **받는분 주소**  
        - `수취인연락처1` → **받는분 전화번호**  
        - `=상품명&옵션정보` → **상품명** (두 값을 그대로 연결)  
        - `수량` → **수량**  
        - `배송메세지` → **메모**  (※ 일부 파일은 `배송메시지` 표기)
        """
    )

st.subheader("스마트스토어 소스 파일 업로드 (키워드 매핑)")
src_file_ss_fixed = st.file_uploader(
    "스마트스토어 형식의 파일 업로드 (예: 스마트스토어.xlsx)",
    type=["xlsx"],
    key="src_smartstore_fixed",
)

run_ss_fixed = st.button("스마트스토어 변환 실행 (키워드 매핑)")
if run_ss_fixed:
    if not src_file_ss_fixed:
        st.error("스마트스토어 소스 파일을 업로드해 주세요.")
    elif tpl_df is None or len(template_columns) == 0:
        st.error("유효한 템플릿이 필요합니다.")
    else:
        try:
            df_ss = read_first_sheet_source_as_text(src_file_ss_fixed)
        except Exception as e:
            st.exception(RuntimeError(f"스마트스토어 소스 파일을 읽는 중 오류: {e}"))
        else:
            def find_col(preferred_names, df):
                norm_cols = {norm_header(c): c for c in df.columns}
                cand_norm = [norm_header(x) for x in preferred_names]
                for n in cand_norm:
                    if n in norm_cols:
                        return norm_cols[n]
                for want in cand_norm:
                    hits = [orig for k, orig in norm_cols.items() if want in k]
                    if hits:
                        return sorted(hits, key=len)[0]
                raise KeyError(f"해당 키워드에 맞는 컬럼을 찾을 수 없습니다: {preferred_names}")

            try:
                col_order = find_col(SS_NAME_MAP["주문번호"], df_ss)
                col_name = find_col(SS_NAME_MAP["받는분 이름"], df_ss)
                col_addr = find_col(SS_NAME_MAP["받는분 주소"], df_ss)
                col_phone = find_col(SS_NAME_MAP["받는분 전화번호"], df_ss)
                col_prod_l = find_col(SS_NAME_MAP["상품명_left"], df_ss)
                col_prod_r = find_col(SS_NAME_MAP["상품명_right"], df_ss)
                col_qty = find_col(SS_NAME_MAP["수량"], df_ss)
                col_memo = find_col(SS_NAME_MAP["메모"], df_ss)
            except Exception as e:
                st.exception(RuntimeError(f"스마트스토어 키워드 매핑 해석 중 오류: {e}"))
            else:
                result_ss = pd.DataFrame(index=range(len(df_ss)), columns=template_columns)

                result_ss["주문번호"] = df_ss[col_order]
                result_ss["받는분 이름"] = df_ss[col_name]
                result_ss["받는분 주소"] = df_ss[col_addr]

                series_phone = df_ss[col_phone].astype(str)
                result_ss["받는분 전화번호"] = series_phone.where(series_phone.str.lower() != "nan", "")

                left_raw = df_ss[col_prod_l].astype(str)
                right_raw = df_ss[col_prod_r].astype(str)
                left = left_raw.where(left_raw.str.lower() != "nan", "")
                right = right_raw.where(right_raw.str.lower() != "nan", "")
                result_ss["상품명"] = (left.fillna("") + right.fillna(""))

                result_ss["수량"] = pd.to_numeric(df_ss[col_qty], errors="coerce")
                result_ss["메모"] = df_ss[col_memo]

                for col in template_columns:
                    if col in tpl_df.columns and tpl_df[col].notna().any():
                        if pd.api.types.is_numeric_dtype(tpl_df[col]) and col != "받는분 전화번호":
                            result_ss[col] = pd.to_numeric(result_ss[col], errors="coerce")

                st.success(f"스마트스토어(키워드) 변환 완료: 총 {len(result_ss)}행")
                st.dataframe(result_ss.head(50))

                out_df_ss = result_ss[template_columns + [c for c in result_ss.columns if c not in template_columns]]
                download_df(
                    out_df_ss,
                    "스마트스토어 변환 결과 다운로드",
                    "스마트스토어 3pl발주용",
                    "ss_conv",
                    sheet_name="발송처리",        # ★ 시트명 고정
                    csv_sep_override=",",         # ★ CSV 쉼표 고정
                    csv_encoding_override=None,   # (필요시 "utf-8-sig"로 고정 가능)
                )

st.markdown("---")

# ======================================================================
# 4) 떠리몰 파일 변환 (키워드 매핑: S&V 규칙)
# ======================================================================
st.markdown("## 떠리몰 파일 변환 (키워드 매핑)")

with st.expander("떠리몰(키워드) → 템플릿 매핑 보기", expanded=False):
    st.markdown(
        """
        **떠리몰 컬럼명(헤더) → 템플릿 컬럼**  
        - `주문번호/주문ID/...` → **주문번호**  
        - `수령자명/받는분/...` → **받는분 이름**  
        - `주소/배송지주소/...` → **받는분 주소**  
        - `수령자연락처/연락처/휴대폰/...` → **받는분 전화번호**  
        - `상품명(S)` & `옵션명:옵션값`(또는 옵션 관련 컬럼) → **상품명**  
          (S와 V가 같으면 V만, 다르면 S+V 그대로 연결)  
        - `수량/구매수/...` → **수량**  
        - `배송메시지/메모/...` → **메모**
        """
    )

st.subheader("떠리몰 소스 파일 업로드 (키워드 매핑)")
src_file_ttarimall = st.file_uploader("떠리몰 형식의 파일 업로드 (예: 떠리몰.xlsx)", type=["xlsx"], key="src_ttarimall")

# 공용 find_col
def find_col(preferred_names, df):
    norm_cols = {norm_header(c): c for c in df.columns}
    cand_norm = [norm_header(x) for x in preferred_names]
    for n in cand_norm:
        if n in norm_cols:
            return norm_cols[n]
    for want in cand_norm:
        hits = [orig for k, orig in norm_cols.items() if want in k]
        if hits:
            return sorted(hits, key=len)[0]
    raise KeyError(f"해당 키워드에 맞는 컬럼을 찾을 수 없습니다: {preferred_names}")

def convert_ttarimall_keywords(df_tm: pd.DataFrame) -> pd.DataFrame:
    col_order = find_col(TTARIMALL_NAME_MAP["주문번호"], df_tm)
    col_name  = find_col(TTARIMALL_NAME_MAP["받는분 이름"], df_tm)
    col_addr  = find_col(TTARIMALL_NAME_MAP["받는분 주소"], df_tm)
    col_phone = find_col(TTARIMALL_NAME_MAP["받는분 전화번호"], df_tm)
    col_s     = find_col(TTARIMALL_NAME_MAP["상품명_S"], df_tm)
    col_v     = find_col(TTARIMALL_NAME_MAP["상품명_V"], df_tm)
    col_qty   = find_col(TTARIMALL_NAME_MAP["수량"], df_tm)
    col_memo  = find_col(TTARIMALL_NAME_MAP["메모"], df_tm)

    result_tm = pd.DataFrame(index=range(len(df_tm)), columns=template_columns)

    result_tm["주문번호"] = df_tm[col_order]
    result_tm["받는분 이름"] = df_tm[col_name]
    result_tm["받는분 주소"] = df_tm[col_addr]

    series_phone = df_tm[col_phone].astype(str)
    result_tm["받는분 전화번호"] = series_phone.where(series_phone.str.lower() != "nan", "")

    s_series_raw = df_tm[col_s].astype(str)
    v_series_raw = df_tm[col_v].astype(str)
    s_series = s_series_raw.where(s_series_raw.str.lower() != "nan", "")
    v_series = v_series_raw.where(v_series_raw.str.lower() != "nan", "")
    same_mask = (s_series == v_series)
    prod_series = v_series.copy()
    prod_series.loc[~same_mask] = s_series[~same_mask] + v_series[~same_mask]
    result_tm["상품명"] = prod_series

    result_tm["수량"] = pd.to_numeric(df_tm[col_qty], errors="coerce")
    result_tm["메모"] = df_tm[col_memo]

    return result_tm

run_ttarimall = st.button("떠리몰 변환 실행 (키워드 매핑)")
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
            try:
                result_tm = convert_ttarimall_keywords(df_tm)

                # 템플릿 숫자형 정렬(전화번호 제외)
                for col in template_columns:
                    if col in tpl_df.columns and tpl_df[col].notna().any():
                        if pd.api.types.is_numeric_dtype(tpl_df[col]) and col != "받는분 전화번호":
                            result_tm[col] = pd.to_numeric(result_tm[col], errors="coerce")

                st.success(f"떠리몰(키워드) 변환 완료: 총 {len(result_tm)}행")
                st.dataframe(result_tm.head(50))

                out_df_tm = result_tm[template_columns + [c for c in result_tm.columns if c not in template_columns]]
                download_df(out_df_tm, "떠리몰 변환 결과 다운로드", "떠리몰 3pl발주용", "ttarimall_conv")
            except Exception as e:
                st.exception(RuntimeError(f"떠리몰 키워드 매핑 해석 중 오류: {e}"))

st.markdown("---")

# ======================================================================
# 5) 배치 처리: 여러 파일 자동 분류 → 일괄 변환 → ZIP 다운로드
# ======================================================================
st.markdown("## 🗂️ 배치 처리 (여러 파일 한번에)")

batch_files = st.file_uploader("여러 엑셀 파일을 한번에 업로드하세요", type=["xlsx"], accept_multiple_files=True, key="batch_files")
run_batch = st.button("배치 변환 실행")

def detect_platform_by_headers(df: pd.DataFrame) -> str:
    headers = [norm_header(c) for c in df.columns]

    def has_any(keys):
        keys_norm = [norm_header(k) for k in keys]
        return any(k in headers for k in keys_norm)

    # 떠리몰 감지 키워드
    if has_any(["수령자명", "수령자연락처", "옵션명:옵션값"]):
        return "TTARIMALL"
    if has_any(["수취인명", "수취인연락처1", "통합배송지"]):
        return "SMARTSTORE"
    if has_any(["최초등록상품명"]) or (has_any(["구매수"]) and has_any(["옵션명"])) or has_any(["배송메시지"]):
        return "COUPANG"
    return "LAORA"

def convert_laora(df_src: pd.DataFrame) -> pd.DataFrame:
    mapping = st.session_state.get("mapping", {})
    if not isinstance(mapping, dict) or not mapping:
        raise RuntimeError("라오라 매핑이 없습니다. 사이드바에서 라오라 매핑을 먼저 저장해 주세요.")
    result = pd.DataFrame(index=range(len(df_src)), columns=template_columns)
    src_cols_by_index = list(df_src.columns)
    resolved_map = {}
    for tpl_header, xl_letters in st.session_state["mapping"].items():
        if not xl_letters:
            continue
        idx = excel_col_to_index(xl_letters)
        if idx >= len(src_cols_by_index):
            raise IndexError(
                f"소스 파일에 {xl_letters} 열(0-based index {idx})이 존재하지 않습니다. "
                f"소스 컬럼 수: {len(src_cols_by_index)}"
            )
        resolved_map[tpl_header] = src_cols_by_index[idx]
    for tpl_header, src_colname in resolved_map.items():
        if tpl_header == "수량":
            result[tpl_header] = pd.to_numeric(df_src[src_colname], errors="coerce")
        elif tpl_header == "받는분 전화번호":
            series = df_src[src_colname].astype(str)
            result[tpl_header] = series.where(series.str.lower() != "nan", "")
        else:
            result[tpl_header] = df_src[src_colname]
    return result

def convert_coupang(df_src: pd.DataFrame) -> pd.DataFrame:
    result = pd.DataFrame(index=range(len(df_src)), columns=template_columns)
    src_cols_by_index = list(df_src.columns)
    resolved_map = {}
    for tpl_header, xl_letters in COUPANG_MAPPING.items():
        idx = excel_col_to_index(xl_letters)
        if idx >= len(src_cols_by_index):
            raise IndexError(
                f"쿠팡 소스에 {xl_letters} 열(0-based index {idx})이 존재하지 않습니다. "
                f"소스 컬럼 수: {len(src_cols_by_index)}"
            )
        resolved_map[tpl_header] = src_cols_by_index[idx]
    for tpl_header, src_colname in resolved_map.items():
        if tpl_header == "수량":
            result[tpl_header] = pd.to_numeric(df_src[src_colname], errors="coerce")
        elif tpl_header == "받는분 전화번호":
            series = df_src[src_colname].astype(str)
            result[tpl_header] = series.where(series.str.lower() != "nan", "")
        else:
            result[tpl_header] = df_src[src_colname]
    return result

def convert_smartstore_keywords(df_ss: pd.DataFrame) -> pd.DataFrame:
    col_order = find_col(SS_NAME_MAP["주문번호"], df_ss)
    col_name = find_col(SS_NAME_MAP["받는분 이름"], df_ss)
    col_addr = find_col(SS_NAME_MAP["받는분 주소"], df_ss)
    col_phone = find_col(SS_NAME_MAP["받는분 전화번호"], df_ss)
    col_prod_l = find_col(SS_NAME_MAP["상품명_left"], df_ss)
    col_prod_r = find_col(SS_NAME_MAP["상품명_right"], df_ss)
    col_qty = find_col(SS_NAME_MAP["수량"], df_ss)
    col_memo = find_col(SS_NAME_MAP["메모"], df_ss)

    result = pd.DataFrame(index=range(len(df_ss)), columns=template_columns)
    result["주문번호"] = df_ss[col_order]
    result["받는분 이름"] = df_ss[col_name]
    result["받는분 주소"] = df_ss[col_addr]
    phone = df_ss[col_phone].astype(str)
    result["받는분 전화번호"] = phone.where(phone.str.lower() != "nan", "")
    lraw = df_ss[col_prod_l].astype(str)
    rraw = df_ss[col_prod_r].astype(str)
    l = lraw.where(lraw.str.lower() != "nan", "")
    r = rraw.where(rraw.str.lower() != "nan", "")
    result["상품명"] = l.fillna("") + r.fillna("")
    result["수량"] = pd.to_numeric(df_ss[col_qty], errors="coerce")
    result["메모"] = df_ss[col_memo]
    return result

def convert_ttarimall_keywords_for_batch(df_tm: pd.DataFrame) -> pd.DataFrame:
    return convert_ttarimall_keywords(df_tm)

def post_numeric_alignment(result_df: pd.DataFrame):
    # 템플릿 숫자형 정렬(전화번호 제외)
    for col in template_columns:
        if col in result_df.columns and col in tpl_df.columns and tpl_df[col].notna().any():
            if pd.api.types.is_numeric_dtype(tpl_df[col]) and col != "받는분 전화번호":
                result_df[col] = pd.to_numeric(result_df[col], errors="coerce")

if run_batch:
    if not batch_files:
        st.error("엑셀 파일을 하나 이상 업로드해 주세요.")
    elif tpl_df is None or len(template_columns) == 0:
        st.error("유효한 템플릿이 필요합니다.")
    else:
        zip_buffer = io.BytesIO()
        logs = []
        with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for f in batch_files:
                fname = getattr(f, "name", "uploaded.xlsx")
                try:
                    df = read_first_sheet_source_as_text(f)
                except Exception as e:
                    logs.append(f"[FAIL] {fname}: 파일 읽기 오류 - {e}")
                    continue

                platform = detect_platform_by_headers(df)
                try:
                    if platform == "TTARIMALL":
                        out_df = convert_ttarimall_keywords_for_batch(df)
                    elif platform == "SMARTSTORE":
                        out_df = convert_smartstore_keywords(df)
                    elif platform == "COUPANG":
                        out_df = convert_coupang(df)
                    else:  # LAORA
                        out_df = convert_laora(df)
                    post_numeric_alignment(out_df)

                    xbuf = io.BytesIO()
                    with pd.ExcelWriter(xbuf, engine="openpyxl") as writer:
                        out_df_sorted = out_df[template_columns + [c for c in out_df.columns if c not in template_columns]]
                        out_df_sorted.to_excel(writer, index=False)
                    base = fname.rsplit(".", 1)[0]
                    out_name = f"{base}__{platform.lower()}_converted.xlsx"
                    zf.writestr(out_name, xbuf.getvalue())

                    logs.append(f"[OK]   {fname}: {platform} → rows={len(out_df)} → {out_name}")
                except Exception as e:
                    logs.append(f"[FAIL] {fname}: {platform} 처리 중 오류 - {e}")

            log_text = "Batch Convert Log - " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n" + "\n".join(logs)
            zf.writestr("batch_convert_log.txt", log_text)

        st.success("배치 변환이 완료되었습니다.")
        st.text_area("변환 로그", value="\n".join(logs), height=200)
        st.download_button(
            label="배치 변환 결과 ZIP 다운로드",
            data=zip_buffer.getvalue(),
            file_name=f"batch_converted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
            mime="application/zip",
        )

st.caption("라오라 / 쿠팡 / 스마트스토어(키워드) / 떠리몰(키워드 S&V) 외 양식도 추가 가능합니다. 규칙만 알려주시면 바로 넣어드릴게요.")

# ======================================================================
# 6) 송장등록: 송장파일(.xls/.xlsx) → 라오/스마트스토어/쿠팡/떠리몰 분류 & 생성
# ======================================================================

# 안전 로더 (.xls/.xlsx)
def _read_excel_any(file, header=0, dtype=str, keep_default_na=False) -> pd.DataFrame:
    """
    안전한 엑셀 로더 (.xlsx/.xls)
      - 업로드 바이트 확보 → BytesIO 로 매 시도마다 새로 읽음
      - .xlsx → openpyxl
      - .xls  → xlrd (권장 버전: 1.2.0)
    """
    name = (getattr(file, "name", "") or "").lower()

    data = None
    if hasattr(file, "getvalue"):
        try:
            data = file.getvalue()
        except Exception:
            data = None
    if data is None:
        try:
            cur = file.tell() if hasattr(file, "tell") else None
            if hasattr(file, "seek"):
                file.seek(0)
            data = file.read()
            if hasattr(file, "seek") and cur is not None:
                file.seek(cur)
        except Exception:
            data = None

    def _read_with(engine: Optional[str]):
        bio = io.BytesIO(data) if data is not None else file
        return pd.read_excel(
            bio, sheet_name=0, header=header, dtype=dtype,
            keep_default_na=keep_default_na, engine=engine,
        )

    try:
        if name.endswith(".xlsx"):
            return _read_with("openpyxl")
        elif name.endswith(".xls"):
            try:
                return _read_with("xlrd")
            except Exception as e:
                raise RuntimeError(
                    "'.xls' 파일을 읽으려면 xlrd가 필요합니다. 권장: pip install \"xlrd==1.2.0\"\n"
                    f"원본 오류: {e}"
                )
        else:
            try:
                return _read_with(None)
            except Exception:
                try:
                    return _read_with("openpyxl")
                except Exception:
                    try:
                        return _read_with("xlrd")
                    except Exception as e:
                        raise RuntimeError(
                            "엑셀 파일을 읽을 수 없습니다. (.xlsx는 openpyxl, .xls는 xlrd 필요)\n"
                            f"원본 오류: {e}"
                        )
    except RuntimeError:
        raise
    except Exception as e:
        raise RuntimeError(f"엑셀 파일을 읽는 중 알 수 없는 오류: {e}")

# 숫자만 남기는 헬퍼 (쿠팡 매칭용)
def _digits_only(x: str) -> str:
    return re.sub(r"\D+", "", str(x or ""))

st.markdown("## 🚚 송장등록")

with st.expander("동작 요약", expanded=False):
    st.markdown(
        """
        - **분류 규칙**
          1) 주문번호에 **`LO`** 포함 → **라스트오더(라오)**
          2) (숫자 기준) **16자리** → **스마트스토어**
        - **라오 출력**: 템플릿 업로드 없이 고정 컬럼  
          **[`주문번호`, `택배사코드(08)`, `송장번호`]**
        - **스마트스토어 출력**: 주문 파일과 **주문번호 매칭** → 송장번호 추가/갱신  
          (결과 **시트명: 발송처리**, `택배사` 기본값=**롯데택배**, 파일명에 타임스탬프)
        - **쿠팡 출력**: **송장파일의 P열(주문번호)** ↔ **쿠팡주문파일의 C열(주문번호)** 를  
          **숫자만 비교**하여 일치 시 **쿠팡주문파일 E열(운송장 번호)** 에 **송장파일의 송장번호** 입력
        - **떠리몰 출력(키워드)**: 떠리몰 주문파일의 **주문번호 컬럼**을 찾아 **송장번호**를 자동 기입  
          (TRACKING_KEYS 중 존재하는 컬럼에 쓰고, 없으면 `송장번호`를 새로 생성)
        """
    )

# 라오 고정 컬럼
LAO_FIXED_TEMPLATE_COLUMNS = ["주문번호", "택배사코드", "송장번호"]

st.subheader("1) 파일 업로드")
invoice_file = st.file_uploader("송장번호 포함 파일 업로드 (예: 송장파일.xls)", type=["xls", "xlsx"], key="inv_file")
ss_order_file = st.file_uploader("스마트스토어 주문 파일 업로드 (선택)", type=["xlsx"], key="inv_ss_orders")
cp_order_file = st.file_uploader("쿠팡 주문 파일 업로드 (선택)", type=["xlsx"], key="inv_cp_orders")
tm_order_file = st.file_uploader("떠리몰 주문 파일 업로드 (선택)", type=["xlsx"], key="inv_tm_orders")

run_invoice = st.button("송장등록 실행")

# 헤더 후보
ORDER_KEYS_INVOICE = ["주문번호", "주문ID", "주문코드", "주문번호1"]
TRACKING_KEYS = ["송장번호", "운송장번호", "운송장", "등기번호", "운송장 번호", "송장번호1"]

SS_ORDER_KEYS = ["주문번호"]
SS_TRACKING_COL_NAME = "송장번호"

# 떠리몰 주문파일에서 주문번호 찾기 후보
TM_ORDER_KEYS = ["주문번호", "주문ID", "주문코드", "주문번호1"]

def build_order_tracking_map(df_invoice: pd.DataFrame):
    """송장파일에서 (주문번호 → 송장번호) 매핑 생성 (헤더명 기반)"""
    order_col = find_col(ORDER_KEYS_INVOICE, df_invoice)
    tracking_col = find_col(TRACKING_KEYS, df_invoice)
    orders = df_invoice[order_col].astype(str)
    tracks = df_invoice[tracking_col].astype(str)
    orders = orders.where(orders.str.lower() != "nan", "")
    tracks = tracks.where(tracks.str.lower() != "nan", "")
    mapping = {}
    for o, t in zip(orders, tracks):
        if o and t:
            mapping[str(o)] = str(t)
    return mapping

def classify_orders(mapping: dict):
    """
    분류:
      - 라오: 'LO' 포함
      - 스마트스토어: 숫자만 16자리
      (쿠팡은 자리수 무시 숫자매칭으로 별도 처리)
    """
    lao, ss = {}, {}
    for o, t in mapping.items():
        s = str(o).strip()
        if "LO" in s.upper():
            lao[s] = t
        elif len(_digits_only(s)) == 16:
            ss[s] = t
    return lao, ss

def make_lao_invoice_df_fixed(lao_map: dict) -> pd.DataFrame:
    """라오 송장: 고정 컬럼으로 DF 생성 (택배사코드=08, 컬럼 순서 고정)"""
    if not lao_map:
        return pd.DataFrame(columns=LAO_FIXED_TEMPLATE_COLUMNS)
    orders = list(lao_map.keys())
    tracks = [lao_map[o] for o in orders]
    out = pd.DataFrame(
        {"주문번호": orders, "택배사코드": ["08"] * len(orders), "송장번호": tracks},
        columns=LAO_FIXED_TEMPLATE_COLUMNS,
    )
    return out

def make_ss_filled_df(ss_map: dict, ss_df: Optional[pd.DataFrame]) -> pd.DataFrame:
    """스마트스토어 주문 파일에 송장번호를 매칭해 추가/갱신 (파일 없으면 2열 매핑만)"""
    if ss_df is None or ss_df.empty:
        if not ss_map:
            return pd.DataFrame()
        df = pd.DataFrame({"주문번호": list(ss_map.keys()), SS_TRACKING_COL_NAME: list(ss_map.values())})
        df["택배사"] = "롯데택배"
        return df

    col_order = find_col(SS_ORDER_KEYS, ss_df)
    out = ss_df.copy()
    if SS_TRACKING_COL_NAME not in out.columns:
        out[SS_TRACKING_COL_NAME] = ""

    existing = out[SS_TRACKING_COL_NAME].astype(str)
    is_empty = (existing.str.lower().eq("nan")) | (existing.str.strip().eq(""))
    mapped = out[col_order].astype(str).map(ss_map).fillna("")
    out.loc[is_empty, SS_TRACKING_COL_NAME] = mapped[is_empty]

    # 택배사 기본값=롯데택배
    if "택배사" not in out.columns:
        out["택배사"] = "롯데택배"
    else:
        ser = out["택배사"].astype(str)
        empty_mask = ser.str.lower().eq("nan") | ser.str.strip().eq("")
        out.loc[empty_mask, "택배사"] = "롯데택배"

    return out

# --- (쿠팡) 송장파일 P열 기반 매핑 생성: 키는 숫자만 ---
def build_inv_map_from_P(df_invoice: pd.DataFrame) -> dict:
    """
    송장파일: P열(주문번호) ↔ 송장번호(여러 헤더명 중 탐색) → {숫자키: 송장번호}
    """
    inv_cols = list(df_invoice.columns)
    try:
        inv_order_col = inv_cols[excel_col_to_index("P")]
    except Exception:
        raise RuntimeError("송장파일에 P열(주문번호)이 없습니다. 송장파일 양식을 확인해 주세요.")
    tracking_col = find_col(TRACKING_KEYS, df_invoice)

    orders = df_invoice[inv_order_col].astype(str).where(lambda s: s.str.lower() != "nan", "")
    tracks = df_invoice[tracking_col].astype(str).where(lambda s: s.str.lower() != "nan", "")

    inv_map = {}
    for o, t in zip(orders, tracks):
        key = _digits_only(o)
        if key and str(t):
            inv_map[key] = str(t)  # 중복 키는 마지막 값 우선
    return inv_map

def make_cp_filled_df_by_letters(df_invoice: Optional[pd.DataFrame],
                                 cp_df: Optional[pd.DataFrame]) -> pd.DataFrame:
    """
    쿠팡 송장등록:
      - 매칭 키: (숫자만 남긴) 송장파일의 P열 주문번호 ↔ (숫자만 남긴) 쿠팡주문파일의 C열 주문번호
      - 쓰기 대상: 쿠팡주문파일의 E열(운송장 번호) ← 송장파일의 '송장번호'
      - 자리수/포맷 무시(숫자만 비교)
    """
    if cp_df is None or cp_df.empty:
        return pd.DataFrame()
    if df_invoice is None or df_invoice.empty:
        return cp_df

    inv_map = build_inv_map_from_P(df_invoice)

    cp_cols = list(cp_df.columns)
    try:
        cp_order_col = cp_cols[excel_col_to_index("C")]  # 매칭 키
    except Exception:
        raise RuntimeError("쿠팡 주문 파일에 C열(주문번호)이 없습니다. 쿠팡 주문파일 양식을 확인해 주세요.")
    try:
        cp_track_col = cp_cols[excel_col_to_index("E")]  # 쓰기 대상
    except Exception:
        cp_track_col = "운송장 번호"
        if cp_track_col not in cp_df.columns:
            cp_df = cp_df.copy()
            cp_df[cp_track_col] = ""
        cp_cols = list(cp_df.columns)

    out = cp_df.copy()
    cp_keys = out[cp_order_col].astype(str).map(_digits_only)
    mapped = cp_keys.map(inv_map)

    mask = mapped.notna() & mapped.astype(str).str.len().gt(0)
    out.loc[mask, cp_track_col] = mapped[mask]

    return out

def make_tm_filled_df(tm_df: Optional[pd.DataFrame], inv_map: dict) -> pd.DataFrame:
    """
    떠리몰 송장등록(키워드):
      - 매칭 키: 떠리몰 주문파일의 '주문번호' (헤더 키워드 탐색)
      - 쓰기 대상: 떠리몰 주문파일의 송장 컬럼 (TRACKING_KEYS 중 존재하는 첫 컬럼, 없으면 '송장번호' 생성)
      - 비교 방식: 문자열 그대로 매칭
    """
    if tm_df is None or tm_df.empty:
        return pd.DataFrame()

    # 1) 주문번호 컬럼 찾기
    tm_order_col = find_col(TM_ORDER_KEYS, tm_df)

    # 2) 송장번호(기입) 컬럼 결정
    tracking_col_candidates = [c for c in TRACKING_KEYS if c in list(tm_df.columns)]
    if tracking_col_candidates:
        tm_tracking_col = tracking_col_candidates[0]
        out = tm_df.copy()
    else:
        tm_tracking_col = "송장번호"
        out = tm_df.copy()
        if tm_tracking_col not in out.columns:
            out[tm_tracking_col] = ""

    # 3) 매핑 적용
    keys = out[tm_order_col].astype(str)
    mapped = keys.map(inv_map)

    mask = mapped.notna() & mapped.astype(str).str.len().gt(0)
    out.loc[mask, tm_tracking_col] = mapped[mask]

    return out

if run_invoice:
    df_invoice = None
    df_ss_orders = None
    df_cp_orders = None
    df_tm_orders = None

    if not invoice_file:
        st.error("송장번호가 포함된 송장파일을 업로드해 주세요. (예: 송장파일.xls)")
    else:
        try:
            df_invoice = _read_excel_any(invoice_file, header=0, dtype=str, keep_default_na=False)
        except Exception as e:
            st.exception(RuntimeError(f"송장파일 읽기 오류: {e}"))
            df_invoice = None

        if ss_order_file:
            try:
                df_ss_orders = read_first_sheet_source_as_text(ss_order_file)
            except Exception as e:
                st.warning(f"스마트스토어 주문 파일을 읽는 중 오류: {e}")
                df_ss_orders = None

        if cp_order_file:
            try:
                df_cp_orders = read_first_sheet_source_as_text(cp_order_file)
            except Exception as e:
                st.warning(f"쿠팡 주문 파일을 읽는 중 오류: {e}")
                df_cp_orders = None

        if tm_order_file:
            try:
                df_tm_orders = read_first_sheet_source_as_text(tm_order_file)
            except Exception as e:
                st.warning(f"떠리몰 주문 파일을 읽는 중 오류: {e}")
                df_tm_orders = None

        if df_invoice is None:
            st.error("송장파일을 읽지 못했습니다. 파일 형식 및 내용(주문번호/송장번호 컬럼)을 확인해 주세요.")
        else:
            try:
                order_track_map = build_order_tracking_map(df_invoice)
                lao_map, ss_map = classify_orders(order_track_map)

                lao_out_df = make_lao_invoice_df_fixed(lao_map)
                ss_out_df = make_ss_filled_df(ss_map, df_ss_orders)
                cp_out_df = make_cp_filled_df_by_letters(df_invoice, df_cp_orders)
                tm_out_df = make_tm_filled_df(df_tm_orders, order_track_map)

                cp_update_cnt = 0
                if df_cp_orders is not None and not df_cp_orders.empty:
                    try:
                        inv_map_tmp = build_inv_map_from_P(df_invoice)
                        cp_cols_tmp = list(df_cp_orders.columns)
                        cp_order_col_tmp = cp_cols_tmp[excel_col_to_index("C")]
                        mapped_tmp = df_cp_orders[cp_order_col_tmp].astype(str).map(_digits_only).map(inv_map_tmp)
                        cp_update_cnt = int((mapped_tmp.notna() & mapped_tmp.astype(str).str.len().gt(0)).sum())
                    except Exception:
                        cp_update_cnt = 0

                # 떠리몰 업데이트 건수 추정
                tm_update_cnt = 0
                if df_tm_orders is not None and not df_tm_orders.empty and tm_out_df is not None and not tm_out_df.empty:
                    try:
                        tm_track_col = next((c for c in TRACKING_KEYS if c in tm_out_df.columns), "송장번호")
                        before = df_tm_orders.get(tm_track_col, pd.Series([""]*len(df_tm_orders))).astype(str).fillna("")
                        after  = tm_out_df.get(tm_track_col, pd.Series([""]*len(tm_out_df))).astype(str).fillna("")
                        tm_update_cnt = int((before != after).sum())
                    except Exception:
                        tm_update_cnt = 0

                st.success(
                    f"분류/매칭 완료: 라오 {len(lao_map)}건 / 스마트스토어 {len(ss_map)}건 / "
                    f"쿠팡 업데이트 예정 {cp_update_cnt}건 / 떠리몰 갱신 {tm_update_cnt}건"
                )
                with st.expander("라오 송장 미리보기", expanded=True):
                    st.dataframe(lao_out_df.head(50))
                with st.expander("스마트스토어 송장 미리보기 (시트명: 발송처리)", expanded=False):
                    st.dataframe(ss_out_df.head(50))
                with st.expander("쿠팡 송장 미리보기", expanded=False):
                    st.dataframe(cp_out_df.head(50))
                with st.expander("떠리몰 송장 미리보기", expanded=False):
                    st.dataframe(tm_out_df.head(50))

                # 다운로드(형식 선택)
                download_df(lao_out_df, "라오 송장 완성 다운로드", "라오 송장 완성", "lao_inv")
                if ss_out_df is not None and not ss_out_df.empty:
                    ss_out_export = ss_out_df.copy()
                    if "택배사" not in ss_out_export.columns:
                        ss_out_export["택배사"] = "롯데택배"
                    else:
                        ser = ss_out_export["택배사"].astype(str)
                        empty_mask = ser.str.lower().eq("nan") | ser.str.strip().eq("")
                        ss_out_export.loc[empty_mask, "택배사"] = "롯데택배"
                    download_df(
                        ss_out_export,
                        "스마트스토어 송장 완성 다운로드",
                        "스마트스토어 송장 완성",
                        "ss_inv",
                        sheet_name="발송처리",      # ★ 시트명 고정
                        csv_sep_override=",",       # ★ CSV 쉼표 고정
                        csv_encoding_override=None,
                    )
                if cp_out_df is not None and not cp_out_df.empty:
                    download_df(cp_out_df, "쿠팡 송장 완성 다운로드", "쿠팡 송장 완성", "cp_inv")
                if tm_out_df is not None and not tm_out_df.empty:
                    download_df(tm_out_df, "떠리몰 송장 완성 다운로드", "떠리몰 송장 완성", "tm_inv")

                if (ss_out_df is None or ss_out_df.empty) and (cp_out_df is None or cp_out_df.empty) and (tm_out_df is None or tm_out_df.empty):
                    st.info("스마트스토어/쿠팡/떠리몰 대상 건이 없거나, 매칭할 주문 파일이 없어 생성 결과가 없습니다.")

            except Exception as e:
                st.exception(RuntimeError(f"송장등록 처리 중 오류: {e}"))
