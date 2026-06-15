import streamlit as st
import pandas as pd
import io
import os
import re

# 상단 메뉴 및 불필요한 UI 숨김 처리
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            header {visibility: hidden;}
            footer {visibility: hidden;}
            .stDeployButton {display:none;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

st.title("BOM 역전개 조회 프로그램")

DEFAULT_MASTER_FILE = "master_bom.csv"
UPLOADED_MASTER_FILE = "uploaded_master_bom.csv"

# --------------------------------------------------
@st.cache_data(show_spinner="초고속 메모리 로딩 중...")
def load_and_process_bom(file_path):
    try:
        try:
            df = pd.read_csv(file_path, dtype=str, encoding='utf-8')
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(file_path, dtype=str, encoding='cp949')
            except UnicodeDecodeError:
                df = pd.read_csv(file_path, dtype=str, encoding='utf-8-sig')
    except Exception as e:
        raise ValueError(f"파일을 읽는 중 오류가 발생했습니다: {e}")
        
    df.columns = df.columns.str.strip()
    required_cols = ['ItemNo', 'PItemNo', 'BOMLevel']
    
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"마스터 데이터에 {missing_cols} 컬럼이 반드시 있어야 합니다.")
        
    for col in df.columns:
        if col in required_cols or '대분류' in str(col) or 'PItemBomRev' in str(col):
            df[col] = df[col].fillna('').astype(str).str.strip()
        
    return df

# --------------------------------------------------
with st.sidebar:
    st.header("⚙️ 마스터 BOM 갱신 (선택)")
    st.write("기본 데이터는 서버에 자동 내장되어 있습니다. 향후 구조가 변경되었을 때만 새 CSV 파일을 업로드하세요.")
    
    master_file = st.file_uploader("마스터 BOM 업로드 (.csv 전용)", type=["csv"], key="master")
    
    if master_file is not None:
        with open(UPLOADED_MASTER_FILE, "wb") as f:
            f.write(master_file.getbuffer())
        
        load_and_process_bom.clear()
        st.success("✅ 새로운 마스터 데이터가 적용되었습니다!")

# --------------------------------------------------
current_file_path = None
if os.path.exists(UPLOADED_MASTER_FILE):
    current_file_path = UPLOADED_MASTER_FILE
elif os.path.exists(DEFAULT_MASTER_FILE):
    current_file_path = DEFAULT_MASTER_FILE

if not current_file_path:
    st.warning("📂 GitHub 저장소에 'master_bom.csv' 파일을 올려두시면 웹사이트가 열릴 때마다 즉시 분석을 완료합니다.")
else:
    try:
        master_df = load_and_process_bom(current_file_path)
            
        st.markdown("### 🔍 단계별 역전개 조회")
        st.write("자재 코드를 입력하여 상위 품목(3번대)을 먼저 확인한 뒤, 원하는 상위 품목의 최종 제품을 조회합니다.")
        
        tab1, tab2 = st.tabs(["📋 텍스트 복사/붙여넣기", "📂 엑셀 파일 업로드"])
        
        codes_from_text = []
        codes_from_file = []
        
        with tab1:
            with st.form("search_form"):
                pasted_text = st.text_area(
                    "조회할 자재 코드를 복사해서 붙여넣으세요. (줄바꿈, 쉼표, 공백 구분 지원)", 
                    height=150,
                    placeholder="예시:\n4AC990533\n4BC010088"
                )
                search_btn = st.form_submit_button("1단계 조회하기 🔍")
            
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
                with st.spinner("⏳ 입력하신 자재의 상위 품목을 역추적하고 있습니다. 잠시만 기다려 주십시오..."):
                    category_col = None
                    for col in master_df.columns:
                        if '대분류' in str(col):
                            category_col = col
                            break
                    
                    master_df_clean = master_df[master_df['BOMLevel'].notna() & (master_df['BOMLevel'] != '')].copy()
                    
                    if 'PItemBomRev' in master_df_clean.columns:
                        master_df_clean['TreeKey'] = master_df_clean['PItemNo'].astype(str) + "_" + \
                                                     master_df_clean['PItemBomRev'].astype(str) + "_" + \
                                                     master_df_clean['BOMLevel'].astype(str)
                    else:
                        master_df_clean['TreeKey'] = master_df_clean['PItemNo'].astype(str) + "_" + \
                                                     master_df_clean['BOMLevel'].astype(str)
                    
                    master_df_unique = master_df_clean.drop_duplicates(subset=['TreeKey'], keep='first')
                    bom_map = master_df_unique.set_index('TreeKey').to_dict('index')
                    
                    all_results = []
                    matched_records = matched_rows.to_dict('records')
                    
                    for row in matched_records:
                        item_no = row.get('ItemNo', '')
                        item_name = row.get('ItemName', '')
                        bom_level = str(row.get('BOMLevel', '')).strip()
                        
                        pitem_no = str(row.get('PItemNo', '')).strip()
                        pitem_name = str(row.get('PItemName', '')).strip()
                        rev = str(row.get('PItemBomRev', '')).strip() if 'PItemBomRev' in master_df_clean.columns else ''
                        prefix = f"{pitem_no}_{rev}_" if 'PItemBomRev' in master_df_clean.columns else f"{pitem_no}_"
                        
                        tokens = bom_level.split('-') if bom_level else []
                        
                        parent_str = ""
                        found_p_item = ""
                        found_p_name = ""
                        found_3_series = False
                        
                        if len(tokens) > 1:
                            for i in range(len(tokens) - 1, 0, -1):
                                p_lvl_str = "-".join(tokens[:i])
                                search_key = prefix + p_lvl_str
                                p_row = bom_map.get(search_key, {})
                                
                                p_item = str(p_row.get('ItemNo', '')).strip()
                                p_name = str(p_row.get('ItemName', '')).strip()
                                
                                if p_item.startswith('3'):
                                    parent_str = f"{p_item}({p_name})"
                                    found_p_item = p_item
                                    found_p_name = p_name
                                    found_3_series = True
                                    break
                        
                        if not found_3_series:
                            continue
                        
                        final_category = row.get(category_col, '') if category_col else ''
                                    
                        # 💡 핵심 수정: 2단계 5칸 분리를 위해 데이터를 각각 분할 저장
                        all_results.append({
                            '입력 자재 코드': item_no,
                            '자재명': item_name if pd.notna(item_name) else "",
                            '상위 품목': parent_str,
                            '상위품목 코드': found_p_item,
                            '상위품목 품명': found_p_name,
                            '최종제품 코드': pitem_no if pd.notna(pitem_no) else "",
                            '최종제품명': pitem_name if pd.notna(pitem_name) else "",
                            '대분류': final_category if pd.notna(final_category) else ""
                        })
                    
                    df_all = pd.DataFrame(all_results).drop_duplicates().reset_index(drop=True)
                
                if df_all.empty:
                    st.warning("입력하신 자재 코드 중, 3번대 상위 품목을 거치는 데이터가 없습니다.")
                else:
                    # ==========================================
                    # 🟢 1단계: 상위 품목 요약표 및 엑셀 다운로드
                    # ==========================================
                    st.markdown("---")
                    st.markdown("#### 🟢 1단계: 투입 자재의 상위 품목 확인")
                    step1_df = df_all[['입력 자재 코드', '자재명', '상위 품목']].drop_duplicates().reset_index(drop=True)
                    st.dataframe(step1_df, hide_index=True)
                    
                    output_step1 = io.BytesIO()
                    with pd.ExcelWriter(output_step1, engine='openpyxl') as writer:
                        step1_df.to_excel(writer, index=False, sheet_name='1단계_요약결과')
                    excel_data_step1 = output_step1.getvalue()

                    st.download_button(
                        label="엑셀로 내려받기 📥",
                        data=excel_data_step1,
                        file_name="BOM_1단계_상위품목_요약.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_step1"
                    )
                    
                    # ==========================================
                    # 🔵 2단계: 체크박스(다중 선택) 및 최종 제품 조회
                    # ==========================================
                    st.markdown("#### 🔵 2단계: 최종 제품 전개")
                    unique_parents = step1_df['상위 품목'].unique().tolist()
                    
                    selected_parents = st.multiselect(
                        "👉 최종 제품을 확인할 상위 품목을 선택하세요 (여러 개 선택 가능):", 
                        options=unique_parents,
                        placeholder="여기를 클릭하여 상위 품목 선택"
                    )
                    
                    if selected_parents:
                        with st.spinner("⏳ 선택하신 상위 품목의 최종 제품 구조를 전개하고 있습니다..."):
                            step2_df = df_all[df_all['상위 품목'].isin(selected_parents)]
                            # 💡 핵심 수정: 화면 및 엑셀 출력 시 정확히 5칸으로 분리하여 산출
                            final_display_df = step2_df[['상위품목 코드', '상위품목 품명', '최종제품 코드', '최종제품명', '대분류']].drop_duplicates().reset_index(drop=True)
                        
                        st.success(f"선택하신 상위 품목이 투입되는 최종 제품 총 {len(final_display_df)}건을 찾았습니다!")
                        st.dataframe(final_display_df, hide_index=True)
                        
                        output_step2 = io.BytesIO()
                        with pd.ExcelWriter(output_step2, engine='openpyxl') as writer:
                            final_display_df.to_excel(writer, index=False, sheet_name='2단계_최종제품결과')
                        excel_data_step2 = output_step2.getvalue()

                        st.download_button(
                            label="엑셀로 내려받기 📥",
                            data=excel_data_step2,
                            file_name="BOM_2단계_최종제품_결과.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_step2"
                        )

    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
