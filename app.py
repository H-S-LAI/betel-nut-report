import streamlit as st
import pandas as pd
import io
import os
from openpyxl import load_workbook

# --- 1. 核心功能：修復亂碼並讀取資料 ---
def load_and_fix_csv(uploaded_file):
    try:
        bytes_data = uploaded_file.getvalue()
        try:
            content = bytes_data.decode('utf-8')
            if '©±' in content: 
                content = content.encode('latin1').decode('big5')
        except:
             content = bytes_data.decode('big5')

        df = pd.read_csv(io.StringIO(content))
        
        col_map = {
            '©±¦W': '店名',
            '«~¦W': '品名',
            '°â¶q': '售量'
        }
        df = df.rename(columns=col_map)
        
        if '店名' in df.columns and '售量' in df.columns:
            df = df[['店名', '品名', '售量']].dropna()
            df['售量'] = pd.to_numeric(df['售量'], errors='coerce').fillna(0)
            return df
        else:
            return None
    except Exception as e:
        st.error(f"檔案 {uploaded_file.name} 讀取失敗: {e}")
        return None

# --- 2. 核心功能：處理 Excel 模板 ---
def fill_excel_template(template_path_or_file, combined_df, grains_per_pack_map):
    # 判斷傳入的是檔案路徑(字串)還是上傳的檔案物件
    if isinstance(template_path_or_file, str):
        wb = load_workbook(template_path_or_file)
    else:
        wb = load_workbook(template_path_or_file)
        
    ws = wb.active

    # 準備數據字典
    data_dict = {}
    for index, row in combined_df.iterrows():
        store = str(row['店名']).strip()
        product = str(row['品名']).strip()
        sales = row['售量']
        
        if store not in data_dict:
            data_dict[store] = {}
        data_dict[store][product] = data_dict[store].get(product, 0) + sales

    # 掃描 header
    header_row = 3
    for r in range(1, 10):
        if ws.cell(row=r, column=1).value == "店名":
            header_row = r
            break
            
    # 掃描品名欄位
    product_col_map = {}
    for col in range(2, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col).value
        if val and isinstance(val, str):
            product_name = val.strip()
            if product_name != "售量" and product_name in grains_per_pack_map:
                product_col_map[product_name] = col

    total_sales_packs = {p: 0 for p in product_col_map}
    row_packs = None
    row_grains = None
    
    # 填寫數據
    for row in range(header_row + 1, ws.max_row + 1):
        cell_val = ws.cell(row=row, column=1).value
        if not cell_val:
            continue
        
        row_label = str(cell_val).strip()
        
        if "銷售包數" in row_label:
            row_packs = row
            continue
        if "銷售粒數" in row_label:
            row_grains = row
            continue
            
        if row_label in data_dict:
            for product, col_idx in product_col_map.items():
                if product in data_dict[row_label]:
                    val = data_dict[row_label][product]
                    ws.cell(row=row, column=col_idx + 1).value = val
                    total_sales_packs[product] += val

    # 填寫統計列
    if row_packs:
        for product, col_idx in product_col_map.items():
            grains_setting = grains_per_pack_map.get(product, 0)
            ws.cell(row=row_packs, column=col_idx).value = grains_setting
            
            total_packs = total_sales_packs.get(product, 0)
            ws.cell(row=row_packs, column=col_idx + 1).value = total_packs

            if row_grains:
                total_grains = total_packs * grains_setting
                ws.cell(row=row_grains, column=col_idx + 1).value = total_grains

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- 3. Streamlit 介面 ---
st.set_page_config(page_title="檳榔報表生成器", layout="wide")
st.title("🏭 檳榔銷售報表自動生成")

# 預設模板檔名
DEFAULT_TEMPLATE = "檳榔銷售統計.xlsx"

# 介面配置
col1, col2 = st.columns([1, 2])

with col1:
    st.markdown("### 1. 模板設定")
    # 檢查預設模板是否存在
    if os.path.exists(DEFAULT_TEMPLATE):
        st.success(f"✅ 已偵測到預設模板：{DEFAULT_TEMPLATE}")
        use_default = st.checkbox("使用預設模板", value=True)
        template_file = DEFAULT_TEMPLATE if use_default else None
        
        if not use_default:
            template_file = st.file_uploader("上傳新模板", type=["xlsx"])
    else:
        st.warning("⚠️ 未偵測到預設模板，請上傳。")
        template_file = st.file_uploader("上傳模板", type=["xlsx"])

with col2:
    st.markdown("### 2. 原始數據")
    source_files = st.file_uploader("請一次上傳所有數據檔案 (特幼, 雙子星...)", type=["csv", "xls"], accept_multiple_files=True)

# 參數設定
default_grains = {
    "特幼": 8, "幼大口": 8, "多粒": 12, "多大口": 12,
    "幼菁": 10, "雙子星": 10, "多菁": 10, "普通": 10
}

st.markdown("### 3. 設定每包粒數")
cols = st.columns(4)
user_grains_setting = {}

for i, (product, default_val) in enumerate(default_grains.items()):
    with cols[i % 4]:
        val = st.number_input(f"{product}", value=default_val, step=1)
        user_grains_setting[product] = val

if st.button("🚀 生成報表", type="primary"):
    if not template_file:
        st.error("找不到模板檔案！請確認已上傳或使用預設模板。")
    elif not source_files:
        st.error("請上傳原始數據檔案。")
    else:
        with st.spinner("處理中..."):
            all_data = []
            for f in source_files:
                df = load_and_fix_csv(f)
                if df is not None:
                    all_data.append(df)
            
            if all_data:
                combined_df = pd.concat(all_data, ignore_index=True)
                try:
                    result_excel = fill_excel_template(template_file, combined_df, user_grains_setting)
                    st.success("完成！")
                    st.download_button(
                        label="📥 下載報表",
                        data=result_excel,
                        file_name="已填寫_檳榔銷售統計.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"錯誤: {e}")
            else:
                st.error("無法讀取數據。")
