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
    st.write("기본 데이터는 서버에 0.1초 만에 켜지도록 자동 내장되어 있습니다. 향후 마스터 데이터가 변경되었을 때만 새 CSV 파일을 업로드하세요.")
    
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
                    placeholder="예시:\n4AC990533\n4BC"
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
                
                # 💡 핵심 수정: 단순히 BOMLevel만 고유키로 쓰면 다른 제품의 트리가 삭제됨.
                # (최종제품코드 + BOM버전 + BOMLevel)을 조합하여 절대 중복되지 않는 고유 키(TreeKey)를 생성.
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
                
                result_data = []
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
                    
                    # 1. 중간 상위 레벨 역추적 (유실 없이 해당 트리의 직계 부모만 정확히 추적)
                    upper_items = []
                    for i in range(len(tokens) - 1, 0, -1):
                        p_lvl_str = "-".join(tokens[:i])
                        search_key = prefix + p_lvl_str
                        p_row = bom_map.get(search_key, {})
                        p_item = p_row.get('ItemNo', '')
                        p_name = p_row.get('ItemName', '')
                        if p_item:
                            upper_items.append(f"{p_item}({p_name})")
                    
                    upper_items_combined = " ➔ ".join(upper_items) if upper_items else "직계 상위 없음"
                    
                    # 2. 최종 제품 정보 (PItemNo) 및 대분류
                    final_pitem_no = pitem_no
                    final_pitem_name = pitem_name
                    final_category = row.get(category_col, '') if category_col else ''
                                
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
                        st.info("💡 화면에는 상위 100건만 표시됩니다. (전체 결과는 엑셀로 다운로드하세요)")
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
