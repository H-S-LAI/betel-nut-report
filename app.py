import streamlit as st
import pandas as pd
import io
import os
from openpyxl import load_workbook

# --- 1. 核心功能：全能讀取與修復 (維持不變，效果很好) ---
def load_and_fix_smart(uploaded_file):
    file_name = uploaded_file.name
    file_ext = os.path.splitext(file_name)[1].lower()
    df = None
    msg = ""

    if file_ext in ['.xls', '.xlsx']:
        try:
            if file_ext == '.xls':
                df = pd.read_excel(uploaded_file, engine='xlrd')
            else:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            msg = "Excel Read Success"
        except Exception as e:
            return None, f"Excel 讀取失敗: {e}"
    else:
        bytes_data = uploaded_file.getvalue()
        content = ""
        decoded_method = ""
        try:
            text_utf8 = bytes_data.decode('utf-8')
            if '©±' in text_utf8 or '§O' in text_utf8: 
                content = text_utf8.encode('latin1', errors='ignore').decode('cp950', errors='ignore')
                decoded_method = "Mojibake Fix"
            else:
                content = text_utf8
                decoded_method = "UTF-8"
        except:
            try:
                content = bytes_data.decode('cp950', errors='ignore')
                decoded_method = "CP950"
            except:
                content = bytes_data.decode('latin1', errors='ignore')
                decoded_method = "Latin1"

        lines = content.splitlines()
        header_row_index = -1
        for i, line in enumerate(lines[:20]): 
            if "店名" in line and "售量" in line:
                header_row_index = i
                break
        if header_row_index == -1:
            return None, f"找不到標題列，請確認檔案內容。"

        try:
            valid_content = "\n".join(lines[header_row_index:])
            df = pd.read_csv(io.StringIO(valid_content))
            msg = "CSV Read Success"
        except Exception as e:
            return None, f"解析 CSV 失敗: {e}"

    if df is not None:
        try:
            target_df = pd.DataFrame()
            df.columns = [str(c).strip() for c in df.columns]

            if '店名' in df.columns and '售量' in df.columns:
                target_df = df
            elif df.shape[1] >= 4:
                target_df = df.iloc[:, [1, 2, 3]].copy()
                target_df.columns = ['店名', '品名', '售量']
            else:
                return None, f"欄位識別失敗。"

            target_df['售量'] = pd.to_numeric(target_df['售量'], errors='coerce').fillna(0)
            target_df = target_df.dropna(subset=['店名'])
            target_df = target_df[target_df['店名'].astype(str).str.contains("店名") == False]
            return target_df, "Success"
        except Exception as e:
            return None, f"資料標準化失敗: {e}"
    return None, "Unknown Error"


# --- 2. 核心功能：填寫 Excel (v8 強力更新版) ---
def fill_excel_template(template_path_or_file, combined_df, grains_per_pack_map):
    if isinstance(template_path_or_file, str):
        wb = load_workbook(template_path_or_file)
    else:
        wb = load_workbook(template_path_or_file)
    ws = wb.active

    # ==========================================
    # 步驟一：先填寫銷售數字 (Data Filling)
    # ==========================================
    # 建立查找字典：Store -> Product -> Sales
    data_dict = {}
    for index, row in combined_df.iterrows():
        store = str(row['店名']).strip()
        product = str(row['品名']).strip()
        sales = row['售量']
        
        if store not in data_dict:
            data_dict[store] = {}
        
        # 模糊匹配產品名稱
        matched_key = product
        for key in grains_per_pack_map.keys():
            if key in product:
                matched_key = key
                break
        data_dict[store][matched_key] = data_dict[store].get(matched_key, 0) + sales

    # 定位 Header (尋找 "店" 開頭的列)
    header_row = 3
    for r in range(1, 10):
        val = ws.cell(row=r, column=1).value
        if val and "店" in str(val):
            header_row = r
            break
            
    # 掃描 Excel 結構 (尋找哪一欄是店名、哪一欄是品名)
    # 這裡我們用一個寬鬆的邏輯：只要該欄位下方填的是 "特幼"，那它就是特幼欄
    
    # 開始填寫數據
    # 為了應對複雜排版，我們掃描所有包含「店名」的欄位
    store_cols = []
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col).value
        if val and "店" in str(val):
            store_cols.append(col)

    # 針對每一區塊 (左、中、右...)
    for store_col in store_cols:
        prod_col = store_col + 1
        sales_col = store_col + 2
        
        # 從 Header 下一行開始往下填
        for r in range(header_row + 1, ws.max_row + 1):
            cell_store = ws.cell(row=r, column=store_col).value
            
            # 遇到 "銷售包數" 就跳過，這不是店名
            if not cell_store or "銷售" in str(cell_store):
                continue
                
            store_name = str(cell_store).strip()
            
            # 取得這一行原本預設的品名 (例如 "特幼")
            cell_prod = ws.cell(row=r, column=prod_col).value
            if not cell_prod:
                continue
            prod_name_in_excel = str(cell_prod).strip()
            
            # 嘗試去 data_dict 找數據
            if store_name in data_dict:
                # 模糊匹配：看 Excel 裡的品名是否包含我們設定的 key
                sales_val = 0
                for key_prod in data_dict[store_name]:
                    if key_prod in prod_name_in_excel or prod_name_in_excel in key_prod:
                        sales_val = data_dict[store_name][key_prod]
                        break
                
                # 填寫銷售量
                if sales_val > 0:
                    ws.cell(row=r, column=sales_col).value = sales_val

    # ==========================================
    # 步驟二：強力更新粒數與總金額 (Green & Blue Cells)
    # ==========================================
    # 策略：掃描 "銷售包數" 的每一列，往上看它是哪個產品，然後更新設定
    
    # 找出所有包含 "銷售包數" 的列 (Row Indices)
    pack_rows = []
    for r in range(1, ws.max_row + 1):
        val = ws.cell(row=r, column=1).value
        if val and "銷售包數" in str(val):
            pack_rows.append(r)

    for r_pack in pack_rows:
        r_grain = -1
        # 找找看下面有沒有 "銷售粒數" (通常在下一行)
        if ws.cell(row=r_pack + 1, column=1).value and "銷售粒數" in str(ws.cell(row=r_pack + 1, column=1).value):
            r_grain = r_pack + 1

        # 掃描這一列的所有欄位
        for col in range(1, ws.max_column + 1):
            # 1. 識別產品：往上看 3 格 (假設數據區有資料)，看看是什麼產品
            # 為了保險，我們往上找直到找到文字
            found_product = None
            for offset in range(1, 5): # 往上找 5 格
                val = ws.cell(row=r_pack - offset, column=col).value
                if val and isinstance(val, str) and len(val) > 1:
                    # 檢查這是不是我們已知的產品名稱
                    for key in grains_per_pack_map.keys():
                        if key in val:
                            found_product = key
                            break
                    if found_product:
                        break
            
            # 2. 如果找到了產品 (例如 "特幼")
            if found_product:
                # 取得使用者設定的粒數 (例如 12)
                setting_val = grains_per_pack_map.get(found_product)
                
                # A. 更新綠色格子 (粒數設定)
                # 位置通常就在這一欄 (r_pack, col)
                ws.cell(row=r_pack, column=col).value = setting_val
                
                # B. 更新藍色格子 (總粒數 = 總包數 * 粒數)
                # 總包數通常在右邊一格 (col + 1)，也就是紅色的格子
                total_packs = ws.cell(row=r_pack, column=col + 1).value
                
                # 確保是數字
                if isinstance(total_packs, (int, float)):
                    total_grains = total_packs * setting_val
                    
                    # 寫入位置：通常在下一列 (r_grain)，且在銷售量那一欄 (col + 1)
                    if r_grain != -1:
                        ws.cell(row=r_grain, column=col + 1).value = total_grains

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- 3. Streamlit 介面 (維持不變) ---
st.set_page_config(page_title="檳榔報表生成器 (v8 強力修復版)", layout="wide")
st.title("🏭 檳榔銷售報表自動生成")

DEFAULT_TEMPLATE = "檳榔銷售統計.xlsx"

col1, col2 = st.columns([1, 2])
with col1:
    st.markdown("### 1. 模板設定")
    if os.path.exists(DEFAULT_TEMPLATE):
        st.success(f"✅ 使用預設模板：{DEFAULT_TEMPLATE}")
        use_default = st.checkbox("使用預設模板", value=True)
        template_file = DEFAULT_TEMPLATE if use_default else None
        if not use_default:
            template_file = st.file_uploader("上傳新模板", type=["xlsx"])
    else:
        st.warning("⚠️ 請上傳 Excel 模板")
        template_file = st.file_uploader("上傳模板", type=["xlsx"])

with col2:
    st.markdown("### 2. 原始數據")
    source_files = st.file_uploader("上傳所有數據檔案 (支援 xls, xlsx, csv)", type=["csv", "xls", "xlsx"], accept_multiple_files=True)

default_grains = {
    "特幼": 8, "幼大口": 8, "多粒": 12, "多大口": 12,
    "幼菁": 10, "雙子星": 10, "多菁": 10, "普通": 10
}

st.markdown("### 3. 設定每包粒數 (將寫入綠色欄位)")
cols = st.columns(4)
user_grains_setting = {}

for i, (product, default_val) in enumerate(default_grains.items()):
    with cols[i % 4]:
        val = st.number_input(f"{product}", value=default_val, step=1)
        user_grains_setting[product] = val

if st.button("🚀 生成報表", type="primary"):
    current_template = template_file if template_file else (DEFAULT_TEMPLATE if os.path.exists(DEFAULT_TEMPLATE) else None)

    if not current_template:
        st.error("找不到模板檔案！")
    elif not source_files:
        st.error("請上傳原始數據檔案。")
    else:
        with st.spinner("正在解析與計算..."):
            all_data = []
            error_logs = []
            
            for f in source_files:
                df, msg = load_and_fix_smart(f)
                if df is not None:
                    all_data.append(df)
                else:
                    error_logs.append(f"❌ {f.name}: {msg}")
            
            if error_logs:
                with st.expander("⚠️ 部分檔案讀取失敗"):
                    for log in error_logs:
                        st.code(log)
            
            if all_data:
                combined_df = pd.concat(all_data, ignore_index=True)
                st.info(f"✅ 成功讀取 {len(combined_df)} 筆資料。")
                
                try:
                    result_excel = fill_excel_template(current_template, combined_df, user_grains_setting)
                    st.success("報表生成成功！設定值已強制更新。")
                    st.download_button(
                        label="📥 下載報表",
                        data=result_excel,
                        file_name="已填寫_檳榔銷售統計.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"填寫 Excel 時發生錯誤: {e}")
            else:
                st.error("沒有任何檔案被成功讀取。")
