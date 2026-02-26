import streamlit as st
import pandas as pd
import os
import io

# 페이지 기본 설정
st.set_page_config(page_title="알러지 양식 변환기", layout="wide")

# ==========================================
# 1. 변환 로직 함수 정의 (나중에 여기에 코드를 채워 넣습니다)
# ==========================================

def logic_cff_83(input_df, template_path):
    """CFF 모드 -> 83 CFF 변환 로직"""
    # TODO: 여기에 실제 변환 코드 작성
    # 임시로 템플릿을 그대로 반환하도록 설정
    return pd.read_excel(template_path)

def logic_cff_26(input_df, template_path):
    """CFF 모드 -> 26 통합 변환 로직"""
    # TODO: 여기에 실제 변환 코드 작성
    return pd.read_excel(template_path)

def logic_hp_83(input_df, template_path):
    """HP 모드 -> 83 HP 변환 로직"""
    # TODO: 여기에 실제 변환 코드 작성
    return pd.read_excel(template_path)

def logic_hp_26(input_df, template_path):
    """HP 모드 -> 26 통합 변환 로직"""
    # TODO: 여기에 실제 변환 코드 작성
    return pd.read_excel(template_path)

# 엑셀 다운로드를 위한 바이너리 변환 함수
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

# ==========================================
# 2. UI 레이아웃 구성
# ==========================================

st.title("📄 알러지 양식 변환기")
st.markdown("---")

# [상단] 입력 및 설정 영역 (2분할)
top_col1, top_col2 = st.columns([1, 1])

with top_col1:
    st.subheader("1. 원본 파일 업로드")
    uploaded_file = st.file_uploader("변환할 엑셀 파일을 올려주세요", type=['xlsx', 'xls'])

with top_col2:
    st.subheader("2. 변환 모드 선택")
    # CFF와 HP를 선택할 수 있는 셀렉트박스
    mode = st.selectbox("업체 타입을 선택하세요", ["CFF", "HP"])
    
    # 선택된 모드에 따라 사용할 템플릿 파일명 미리 지정
    if mode == "CFF":
        st.info("💡 [CFF 모드] '83 CFF' 및 '26 통합' 양식으로 변환합니다.")
    else:
        st.info("💡 [HP 모드] '83 HP' 및 '26 통합' 양식으로 변환합니다.")

st.markdown("---")

# [하단] 실행 및 결과 영역 (2분할)
btm_col1, btm_col2 = st.columns([1, 1])

# 결과물을 담을 변수 초기화 (세션 스테이트 사용)
if 'result_83' not in st.session_state:
    st.session_state.result_83 = None
if 'result_26' not in st.session_state:
    st.session_state.result_26 = None

with btm_col1:
    st.subheader("3. 변환 실행")
    if st.button("변환 시작", type="primary", use_container_width=True):
        if uploaded_file is not None:
            try:
                # 원본 읽기
                input_df = pd.read_excel(uploaded_file)
                
                # 템플릿 경로 설정 (상대 경로)
                base_path = "template"
                
                if mode == "CFF":
                    # CFF 로직 실행
                    res_83 = logic_cff_83(input_df, os.path.join(base_path, "83 CFF.xlsx"))
                    res_26 = logic_cff_26(input_df, os.path.join(base_path, "26 통합.xlsx"))
                else:
                    # HP 로직 실행
                    res_83 = logic_hp_83(input_df, os.path.join(base_path, "83 HP.xlsx"))
                    res_26 = logic_hp_26(input_df, os.path.join(base_path, "26 통합.xlsx"))
                
                # 결과를 세션에 저장 (화면이 리로딩돼도 다운로드 버튼 유지)
                st.session_state.result_83 = to_excel(res_83)
                st.session_state.result_26 = to_excel(res_26)
                
                st.success("변환이 완료되었습니다! 오른쪽에서 다운로드하세요. 👉")
                
            except Exception as e:
                st.error(f"오류가 발생했습니다: {e}")
        else:
            st.warning("먼저 원본 파일을 업로드해주세요.")

with btm_col2:
    st.subheader("4. 결과물 다운로드")
    
    if st.session_state.result_83 and st.session_state.result_26:
        # 파일명 접두사 설정
        prefix = "CFF" if mode == "CFF" else "HP"
        
        # 다운로드 버튼 1: 83 양식
        st.download_button(
            label=f"📥 {prefix}_83 양식 다운로드",
            data=st.session_state.result_83,
            file_name=f"{prefix}_83_Converted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        # 다운로드 버튼 2: 26 통합 양식
        st.download_button(
            label=f"📥 {prefix}_26 통합 다운로드",
            data=st.session_state.result_26,
            file_name=f"{prefix}_26_Converted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.write("왼쪽에서 '변환 시작' 버튼을 누르면 다운로드 버튼이 나타납니다.")
