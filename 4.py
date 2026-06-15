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
    
    # 💡 속도 최적화 1: 전체 열이 아닌 필수 핵심 열만 선별하여 공백 제거 연산 수행
    df.columns = df.columns.str.strip()
    required_cols = ['ItemNo', 'PItemNo', 'BOMLevel']
    
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"마스터 BOM 엑셀 파일에 {missing_cols} 컬럼이 반드시 있어야 합니다.")
        
    for col in df.columns:
        if col in required_cols or '대분류' in str(col):
            df[col] = df[col].fillna('').astype(str).str.strip()
        
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
            
        st.markdown("### 🔍 다계층 BOM 역전개 조회")
        st.write("조회할 자재 코드(ItemNo)를 입력하거나 엑셀 파일을 업로드하세요. 계층 구조를 역추적합니다.")
        
        tab1, tab2 = st.tabs(["📋 텍스트 복사/붙여넣기", "📂 엑셀 파일 업로드"])
        
        codes_from_text = []
        codes_from_file = []
        
        with tab1:
            with st.form("search_form"):
                pasted_text = st.text_area(
                    "조회할 자재 코드를 복사해서 붙여넣으세요. (줄바꿈, 쉼표, 공백 구분 지원)", 
                    height=150,
                    placeholder="예시:\n4AC\n4BC"
                )
                search_btn = st.form_submit_button("조회하기 🔍")
            
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
            matched_rows = master_df[master_df['ItemNo'].isin(target_codes)]
            
            if matched_rows.empty:
                st.warning("마스터 BOM에서 일치하는 자재 코드를 찾지 못했습니다.")
            else:
                category_col = None
                for col in master_df.columns:
                    if '대분류' in str(col):
                        category_col = col
                        break
                
                # 💡 속도 최적화 2: 수만 행을 하나씩 순회하는 iterrows()를 폐기하고, 메모리 상에서 index 딕셔너리로 초고속 직렬화
                master_df_clean = master_df[master_df['BOMLevel'].notna() & (master_df['BOMLevel'] != '')]
                bom_map = master_df_clean.set_index('BOMLevel').to_dict('index')
                
                result_data = []
                matched_records = matched_rows.to_dict('records')
                
                for row in matched_records:
                    item_no = row.get('ItemNo', '')
                    item_name = row.get('ItemName', '')
                    bom_level = str(row.get('BOMLevel', '')).strip()
                    
                    tokens = bom_level.split('-') if bom_level else []
                    
                    # 1. 중간 상위 레벨 역추적
                    upper_items = []
                    for i in range(len(tokens) - 1, 1, -1):
                        p_lvl_str = "-".join(tokens[:i])
                        p_row = bom_map.get(p_lvl_str, {})
                        p_item = p_row.get('ItemNo', '')
                        p_name = p_row.get('ItemName', '')
                        if p_item:
                            upper_items.append(f"{p_item}({p_name})")
                    
                    upper_items_combined = " ➔ ".join(upper_items) if upper_items else "직계 상위 없음"
                    
                    # 2. 최종 상위 제품 정보 및 대분류 추출
                    final_pitem_no = ""
                    final_pitem_name = ""
                    final_category = ""
                    
                    if len(tokens) >= 2:
                        level_2_str = "-".join(tokens[:2])
                        level_2_row = bom_map.get(level_2_str, {})
                        final_pitem_no = level_2_row.get('PItemNo', '')
                        final_pitem_name = level_2_row.get('PItemName', '')
                        if category_col:
                            final_category = level_2_row.get(category_col, '')
                    else:
                        final_pitem_no = row.get('PItemNo', '')
                        final_pitem_name = row.get('PItemName', '')
                        if category_col:
                            final_category = row.get(category_col, '')
                    
                    if not final_category or str(final_category).lower() == 'nan':
                        if category_col:
                            final_category = row.get(category_col, '')
                            if not final_category or str(final_category).lower() == 'nan':
                                final_category = ""
                                
                    result_data.append({
                        '입력 자재 코드': item_no,
                        '자재명': item_name if pd.notna(item_name) else "",
                        '본인 BOM 레벨': bom_level,
                        '중간 상위 계층 목록': upper_items_combined,
                        '최종 제품 코드(PItemNo)': final_pitem_no if pd.notna(final_pitem_no) else "",
                        '최종 제품명(PItemName)': final_pitem_name if pd.notna(final_pitem_name) else "",
                        '대분류': final_category if pd.notna(final_category) else ""
                    })
                
                final_df = pd.DataFrame(result_data).drop_duplicates().reset_index(drop=True)
                
                if final_df.empty:
                    st.warning("역전개 조건을 만족하는 데이터 조합을 만들지 못했습니다.")
                else:
                    st.success(f"총 {len(final_df)}건의 계층 역전개 결과를 산출했습니다!")
                    
                    if len(final_df) > 100:
                        st.info("💡 노트북 과부하를 막기 위해 화면에는 상위 100건만 미리보기로 표시됩니다. (전체 결과는 엑셀로 다운로드하세요)")
                        st.dataframe(final_df.head(100), hide_index=True)
                    else:
                        st.dataframe(final_df, hide_index=True)

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_df.to_excel(writer, index=False, sheet_name='다계층 역전개 결과')
                    excel_data = output.getvalue()

                    st.download_button(
                        label="조회 결과 엑셀로 내려받기 📥",
                        data=excel_data,
                        file_name="BOM_다계층_역전개_결과.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
