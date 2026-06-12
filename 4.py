import streamlit as st
import pandas as pd
import io
import os
import re

st.title("BOM 역전개 조회 프로그램")

SAVE_FILE_PATH = "saved_master_bom.xlsx"

# --------------------------------------------------
@st.cache_data(show_spinner="마스터 BOM을 분석 중입니다.")
def load_and_process_bom(file_path):
    df = pd.read_excel(file_path, dtype=str)
    
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()
    
    if 'ItemNo' not in df.columns or 'PItemNo' not in df.columns:
        raise ValueError("마스터 BOM 엑셀 파일에 'ItemNo'와 'PItemNo' 컬럼이 반드시 있어야 합니다.")
        
    return df

# --------------------------------------------------
with st.sidebar:
    st.header("⚙️ 마스터 BOM 관리")
    st.write("BOM 구조가 변경되었을 때만 갱신하세요.")
    
    master_file = st.file_uploader("마스터 BOM 엑셀 업로드", type=["xlsx"], key="master")
    
    if master_file is not None:
        with open(SAVE_FILE_PATH, "wb") as f:
            f.write(master_file.getbuffer())
        
        load_and_process_bom.clear()
        st.success("✅ 마스터 BOM이 성공적으로 갱신되었습니다!")

# --------------------------------------------------
if not os.path.exists(SAVE_FILE_PATH):
    st.warning("📂 왼쪽 메뉴에서 기준이 되는 [마스터 BOM]을 먼저 업로드해 주세요.")
else:
    try:
        master_df = load_and_process_bom(SAVE_FILE_PATH)
            
        st.markdown("### 🔍 단가 변동 자재 조회 (모든 품목 대상)")
        st.write("조회할 자재 코드를 입력하거나 엑셀 파일을 업로드하세요. 상위 품목을 산출합니다.")
        
        tab1, tab2 = st.tabs(["📋 텍스트 복사/붙여넣기", "📂 엑셀 파일 업로드"])
        
        codes_from_text = []
        codes_from_file = []
        
        with tab1:
            # 💡 핵심 수정: Form을 생성하여 버튼을 누르기 전까지는 코드가 자동 실행되지 않도록 통제함
            with st.form("search_form"):
                pasted_text = st.text_area(
                    "조회할 자재 코드를 복사해서 붙여넣으세요. (줄바꿈, 쉼표, 공백 구분 지원)", 
                    height=150,
                    placeholder="예시:\n40001\n40002\n40003"
                )
                search_btn = st.form_submit_button("조회하기 🔍")
            
            # 입력된 텍스트가 존재하면 코드 분리 작업 수행
            if pasted_text.strip():
                raw_codes = re.split(r'[\n,\s]+', pasted_text)
                codes_from_text = [str(x).strip() for x in raw_codes if str(x).strip() != '']
                
        with tab2:
            target_file = st.file_uploader("조회 대상 엑셀 파일 업로드", type=["xlsx"], key="target")
            if target_file is not None:
                target_df = pd.read_excel(target_file, dtype=str, header=None)
                all_values = target_df.values.flatten().tolist()
                codes_from_file = [str(x).strip() for x in all_values if str(x).strip() != 'nan' and str(x).strip() != '']

        target_codes = list(set(codes_from_text + codes_from_file))
        
        if not target_codes:
            st.info("💡 위의 입력창에 코드를 붙여넣고 [조회하기]를 누르거나, 엑셀 파일을 업로드해 주세요.")
        else:
            filtered_df = master_df[master_df['ItemNo'].isin(target_codes)]
            
            target_columns = ['ItemNo', 'ItemName', 'ItemNo 거래처', 'PItemNo', 'PItemName', 'PItemNo 거래처']
            
            display_columns = []
            for col in target_columns:
                if col in filtered_df.columns:
                    display_columns.append(col)
            
            final_df = filtered_df[display_columns].drop_duplicates().reset_index(drop=True)
            
            rename_dict = {
                'ItemNo': '투입 자재 코드',
                'ItemName': '투입 자재명',
                'ItemNo 거래처': '투입자재 거래처',
                'PItemNo': '상위 품목 코드',
                'PItemName': '상위 품목명',
                'PItemNo 거래처': '상위품목 거래처'
            }
            final_df = final_df.rename(columns=rename_dict)
            
            if final_df.empty:
                st.warning("일치하는 투입 자재 코드를 찾았으나, 연결된 상위 품목 데이터가 없습니다.")
            else:
                st.success(f"총 {len(final_df)}건의 매칭 결과를 산출했습니다!")
                
                if len(final_df) > 100:
                    st.info("💡 노트북 과부하를 막기 위해 화면에는 상위 100건만 미리보기로 표시됩니다. (전체 결과는 엑셀로 다운로드하세요)")
                    st.dataframe(final_df.head(100), hide_index=True)
                else:
                    st.dataframe(final_df, hide_index=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='필터링 결과')
                excel_data = output.getvalue()

                st.download_button(
                    label="조회 결과 엑셀로 내려받기 📥",
                    data=excel_data,
                    file_name="단가변동_영향도_결과.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
