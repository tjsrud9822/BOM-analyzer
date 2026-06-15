import streamlit as st
import pandas as pd
import io
import os
import re

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
        st.write("자재 코드(ItemNo)를 입력하여 상위 품목(3번대)을 먼저 확인한 뒤, 원하는 상위 품목의 최종 제품을 조회합니다.")
        
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
                    
                    # 💡 핵심 수정: 무조건 1계층 위를 찾는 것이 아니라, 
                    # 트리를 거슬러 올라가며 '3'으로 시작하는 품목이 나올 때까지 탐색합니다.
                    parent_str = "3번대 상위품목 없음 (5번 등에 직투입)"
                    if len(tokens) > 1:
                        for i in range(len(tokens) - 1, 0, -1):
                            p_lvl_str = "-".join(tokens[:i])
                            search_key = prefix + p_lvl_str
                            p_row = bom_map.get(search_key, {})
                            
                            p_item = str(p_row.get('ItemNo', '')).strip()
                            p_name = str(p_row.get('ItemName', '')).strip()
                            
                            # 5번 등을 패스하고, 3으로 시작하는 코드를 발견하는 순간 고정하고 멈춤
                            if p_item.startswith('3'):
                                parent_str = f"{p_item}({p_name})"
                                break
                    
                    final_pitem_combined = f"{pitem_no}({pitem_name})" if pitem_no else ""
                    final_category = row.get(category_col, '') if category_col else ''
                                
                    all_results.append({
                        '입력 자재 코드': item_no,
                        '자재명': item_name if pd.notna(item_name) else "",
                        '상위 품목': parent_str,
                        '최종 제품(코드+제품명)': final_pitem_combined,
                        '대분류': final_category if pd.notna(final_category) else ""
                    })
                
                df_all = pd.DataFrame(all_results).drop_duplicates().reset_index(drop=True)
                
                # --- 1단계: 상위 품목 요약표 ---
                st.markdown("---")
                st.markdown("#### 🟢 1단계: 투입 자재의 상위 품목 확인")
                step1_df = df_all[['입력 자재 코드', '자재명', '상위 품목']].drop_duplicates().reset_index(drop=True)
                st.dataframe(step1_df, hide_index=True)
                
                # --- 2단계: 체크박스(다중 선택) 및 최종 제품 조회 ---
                st.markdown("#### 🔵 2단계: 최종 제품 전개")
                unique_parents = step1_df['상위 품목'].unique().tolist()
                
                selected_parents = st.multiselect(
                    "👉 최종 제품을 확인할 상위 품목을 선택하세요 (여러 개 선택 가능):", 
                    options=unique_parents,
                    placeholder="여기를 클릭하여 상위 품목 선택"
                )
                
                if selected_parents:
                    step2_df = df_all[df_all['상위 품목'].isin(selected_parents)]
                    final_display_df = step2_df[['상위 품목', '최종 제품(코드+제품명)', '대분류']].drop_duplicates().reset_index(drop=True)
                    
                    st.success(f"선택하신 상위 품목이 투입되는 최종 제품 총 {len(final_display_df)}건을 찾았습니다!")
                    st.dataframe(final_display_df, hide_index=True)
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_display_df.to_excel(writer, index=False, sheet_name='최종제품_역전개_결과')
                    excel_data = output.getvalue()

                    st.download_button(
                        label="선택 결과 엑셀로 내려받기 📥",
                        data=excel_data,
                        file_name="BOM_최종제품_역전개_결과.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
