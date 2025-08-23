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
st.caption("템플릿 업로드 없이도 기본 템플릿으로 변환하고, 화면에서 매핑 규칙을 자유롭게 수정할 수 있습니다.")

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
    """소스(1.xlsx)는 전 컬럼을 문자열로 읽어 전화번호 앞 0 보존"""
    # keep_default_na=False: 빈 셀을 NaN이 아닌 빈 문자열로 유지 (전화번호 'nan' 문자열화 방지)
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
DEFAULT_MAPPING = {
    "주문번호": "D",
    "받는분 이름": "AM",
    "받는분 주소": "AP",
    "받는분 전화번호": "AN",
    "상품명": "S",
    "수량": "V",
    "메모": "AR",
}

# -------------------------- Sidebar --------------------------
st.sidebar.header("옵션")
use_uploaded_template = st.sidebar.checkbox("템플릿(2.xlsx) 직접 업로드", value=False)
max_letter_cols = st.sidebar.slider(
    "최대 열 범위(Excel 문자)",
    min_value=52,
    max_value=156,
    value=104,
    step=26,
    help="드롭다운에 표시할 엑셀 열 문자 개수",
)
st.sidebar.divider()
st.sidebar.subheader("매핑 저장/불러오기")
mapping_upload = st.sidebar.file_uploader("매핑 JSON 불러오기", type=["json"], key="mapping_json")
prepare_download = st.sidebar.button("현재 매핑 JSON 다운로드 준비")

# -------------------------- 1) 업로드 섹션 (항상 보이게) --------------------------
st.subheader("파일 업로드")
src_file = st.file_uploader("1과 같은 양식의 파일 업로드 (예: 1.xlsx)", type=["xlsx"], key="src")

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
current_mapping = ensure_mapping_initialized(template_columns, DEFAULT_MAPPING)

# -------------------------- 2) 매핑 에디터 --------------------------
st.subheader("매핑 규칙 편집")
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
        st.success("매핑 JSON을 불러왔습니다.")
    except Exception as e:
        st.warning(f"매핑 JSON 불러오기 실패: {e}")

edited_mapping = {}
with st.form("mapping_form"):
    for col in template_columns:
        default_val = current_mapping.get(col, "")
        if default_val not in letters:
            default_val = ""
        options = [""] + letters
        sel = st.selectbox(
            f"{col} ⟶ 1.xlsx 열 문자 선택",
            options=options,
            index=(options.index(default_val) if default_val in options else 0),
            key=f"map_{col}",
        )
        edited_mapping[col] = sel
    if st.form_submit_button("매핑 저장"):
        st.session_state["mapping"] = {k: v for k, v in edited_mapping.items() if v}
        current_mapping = st.session_state["mapping"]
        st.success("매핑을 저장했습니다.")

if prepare_download:
    mapping_bytes = json.dumps(current_mapping, ensure_ascii=False, indent=2).encode("utf-8")
    st.download_button(
        label="현재 매핑 JSON 다운로드",
        data=mapping_bytes,
        file_name="mapping.json",
        mime="application/json",
    )

st.divider()
run = st.button("변환 실행")

# -------------------------- 3) 변환 실행 --------------------------
if run:
    if not src_file:
        st.error("소스 파일(1.xlsx)을 업로드해 주세요.")
    elif tpl_df is None or len(template_columns) == 0:
        st.error("유효한 템플릿이 필요합니다.")
    else:
        try:
            # ✅ 전화번호 0 보존: 소스는 무조건 문자열로 읽음
            df_src = read_first_sheet_source_as_text(src_file)
        except Exception as e:
            st.exception(RuntimeError(f"소스 파일을 읽는 중 오류가 발생했습니다: {e}"))
        else:
            result = pd.DataFrame(index=range(len(df_src)), columns=template_columns)

            mapping = st.session_state.get("mapping", {})
            if not isinstance(mapping, dict) or not mapping:
                st.error("저장된 매핑이 없습니다. 매핑 규칙을 먼저 설정해 주세요.")
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
                    st.exception(RuntimeError(f"매핑 인덱스 계산 중 오류: {e}"))
                else:
                    for tpl_header, src_colname in resolved_map.items():
                        try:
                            if tpl_header == "수량":
                                # 수량만 숫자로 변환
                                result[tpl_header] = pd.to_numeric(df_src[src_colname], errors="coerce")
                            elif tpl_header == "받는분 전화번호":
                                # 전화번호는 항상 문자열(앞 0 보존). 빈값/NaN은 빈 문자열 처리
                                series = df_src[src_colname].astype(str)
                                # pandas가 공백/NaN을 'nan'으로 캐스팅한 경우 제거
                                result[tpl_header] = series.where(series.str.lower() != "nan", "")
                            else:
                                result[tpl_header] = df_src[src_colname]
                        except KeyError:
                            st.warning(f"소스 컬럼 '{src_colname}'(매핑: {tpl_header})을(를) 찾을 수 없습니다. 해당 필드는 비워집니다.")

                    # 템플릿의 숫자형 샘플이 있으면 타입 정렬(선택)
                    for col in template_columns:
                        if col in tpl_df.columns and tpl_df[col].notna().any():
                            if pd.api.types.is_numeric_dtype(tpl_df[col]) and col != "받는분 전화번호":
                                result[col] = pd.to_numeric(result[col], errors="coerce")

                    st.success(f"변환 완료: 총 {len(result)}행")
                    st.dataframe(result.head(50))

                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                        out_df = result[template_columns + [c for c in result.columns if c not in template_columns]]
                        out_df.to_excel(writer, index=False)
                    st.download_button(
                        label="변환 결과 다운로드 (output.xlsx)",
                        data=buffer.getvalue(),
                        file_name="output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

st.markdown("---")
st.caption("전화번호/주소 정규화, 고정값 채우기, 열 머리글 기반 매핑 등도 원하시면 바로 추가해 드릴게요.")
