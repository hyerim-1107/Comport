import streamlit as st
import pandas as pd
from io import BytesIO
import os

st.title("엑셀 중복 제거")

uploaded_file = st.file_uploader("엑셀 파일을 선택하세요 (.xlsx)", type=["xlsx"])
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.write("업로드한 파일 미리보기", df.head())
    except Exception as e:
        st.error(f"파일 읽는 중 오류가 발생했습니다:\n{e}")
    else:
        # "전화번호" 키워드가 포함된 컬럼 찾기
        phone_cols = [col for col in df.columns if "전화번호" in col]
        if not phone_cols:
            st.error("파일에 '전화번호' 키워드가 포함된 컬럼이 없습니다.")
        else:
            # 중복 제거: 찾은 첫 번째 '전화번호' 관련 컬럼을 기준으로 마지막 데이터만 유지
            df_cleaned = df.drop_duplicates(subset=[phone_cols[0]], keep="last")
            
            # 업로드된 파일 이름에서 출력 파일명 생성 (예: ABC_cleaned.xlsx)
            base, ext = os.path.splitext(uploaded_file.name)
            output_filename = base + "_cleaned" + ext

            # DataFrame을 메모리 내 Excel 파일로 저장
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_cleaned.to_excel(writer, index=False)
            output.seek(0)

            st.download_button(
                label="정리된 파일 다운로드",
                data=output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
