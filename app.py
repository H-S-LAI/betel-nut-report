import streamlit as st
import pandas as pd
import io
import os
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

# --- 1. 核心功能：全能讀取與修復 (維持不變) ---
def load_and_fix_smart(uploaded_file):
    file_name = uploaded_file.name
    file_ext = os.path.splitext(file_name)[1].lower()
    df = None

    if file_ext in ['.xls', '.xlsx']:
        try:
            if file_ext == '.xls':
                df = pd.read_excel(uploaded_file, engine='xlrd')
            else:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
        except Exception as e:
            return None, f"Excel 讀取失敗: {e}"
    else:
        bytes_data = uploaded_file.getvalue()
        content = ""
        try:
            text_utf8 = bytes_data.decode('utf-8')
            if '©±' in text_utf8 or '§O' in text_utf8: 
                content = text_utf8.encode('latin1', errors='ignore').decode('cp950', errors='ignore')
            else:
                content = text_utf8
        except:
            try:
                content = bytes_data.decode('cp950', errors='ignore')
            except:
                content = bytes_data.decode('latin1', errors='ignore')

        lines = content.splitlines()
        header_row_index = -1
        for i, line in enumerate(lines[:20]): 
            if "店名" in line and "售量" in line:
                header_row_index = i
                break
        if header_row_index == -1:
            return None, f"找不到標題列。"

        try:
            valid_content = "\n".join(lines[header_row_index:])
            df = pd.read_csv(io.StringIO(valid_content))
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

# --- 功能：安全寫入 (解決 MergedCell 錯誤) ---
def safe_write(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                top_left = ws.cell(row=rng.min_row, column=rng.min_col)
                top_left.value = value
                return
    else:
        cell.value = value

# --- 2. 核心功能：填寫 Excel (V11 智慧跳躍版) ---
def fill_excel_template(template_path_or_file, combined_df, grains_per_pack_map):
    if isinstance(template_path_or_file, str):
        wb = load_workbook(template_path_or_file)
    else:
        wb = load_workbook(template_path_or_file)
    ws = wb.active

    # ==========================================
    # 準備工作
    # ==========================================
    global_total_grains_by_product = {} 
    global_total_packs_all = 0
    col_product_map = {}

    # 1. 整理銷售數據
    data_dict = {}
    for index, row in combined_df.iterrows():
        store = str(row['店名']).strip()
        product = str(row['品名']).strip()
        sales = row['售量']
        
        if store not in data_dict:
            data_dict[store] = {}
        
        matched_key = product
        for key in grains_per_pack_map.keys():
            if key in product:
                matched_key = key
                break
        data_dict[store][matched_key] = data_dict[store].get(matched_key, 0) + sales

    # 定位 Header
    header_row = 3
    for r in range(1, 10):
        val = ws.cell(row=r, column=1).value
        if val and "店" in str(val):
            header_row = r
            break
            
    # 2. 填寫分店數據
    store_cols = []
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col).value
        if val and "店" in str(val):
            store_cols.append(col)

    for store_col in store_cols:
        prod_col = store_col + 1
        sales_col = store_col + 2
        
        for r in range(header_row + 1, ws.max_row + 1):
            cell_store = ws.cell(row=r, column=store_col).value
            if not cell_store or "銷售" in str(cell_store):
                continue
            
            store_name = str(cell_store).strip()
            cell_prod = ws.cell(row=r, column=prod_col).value
            if not cell_prod:
                continue
            prod_name_in_excel = str(cell_prod).strip()
            
            if store_name in data_dict:
                sales_val = 0
                for key_prod in data_dict[store_name]:
                    if key_prod in prod_name_in_excel or prod_name_in_excel in key_prod:
                        sales_val = data_dict[store_name][key_prod]
                        break
                if sales_val > 0:
                    safe_write(ws, r, sales_col, sales_val)

    # ==========================================
    # 3. 處理「紅色包數」與「藍色粒數」
    # ==========================================
    pack_rows = []
    for r in range(1, ws.max_row + 1):
        val = ws.cell(row=r, column=1).value
        if val and "銷售包數" in str(val):
            pack_rows.append(r)

    for r_pack in pack_rows:
        r_grain = -1
        if ws.cell(row=r_pack + 1, column=1).value and "銷售粒數" in str(ws.cell(row=r_pack + 1, column=1).value):
            r_grain = r_pack + 1

        for col in range(1, ws.max_column + 1):
            found_product = None
            for offset in range(1, 6):
                val = ws.cell(row=r_pack - offset, column=col).value
                if val and isinstance(val, str) and len(val) > 1:
                    for key in grains_per_pack_map.keys():
                        if key in val:
                            found_product = key
                            break
                    if found_product:
                        break
            
            if found_product:
                col_product_map[col] = found_product
                col_product_map[col + 1] = found_product

                # 更新綠色 (粒數設定)
                setting_val = grains_per_pack_map.get(found_product)
                safe_write(ws, r_pack, col, setting_val)
                
                # 計算紅色
                current_red_sum = 0
                for offset in range(1, 20):
                    r_scan = r_pack - offset
                    if r_scan <= header_row: break
                    val = ws.cell(row=r_scan, column=col + 1).value
                    if isinstance(val, (int, float)):
                        current_red_sum += val
                
                # 寫入紅色
                cell_red = ws.cell(row=r_pack, column=col + 1)
                cell_red_val = 0
                if isinstance(cell_red, MergedCell):
                     cell_red_val = current_red_sum
                else:
                     cell_red_val = cell_red.value
                     if not isinstance(cell_red_val, (int, float)):
                         cell_red_val = current_red_sum
                
                safe_write(ws, r_pack, col + 1, cell_red_val)
                global_total_packs_all += cell_red_val

                # 寫入藍色
                total_grains = cell_red_val * setting_val
                if r_grain != -1:
                    safe_write(ws, r_grain, col + 1, total_grains)
                
                if found_product not in global_total_grains_by_product:
                    global_total_grains_by_product[found_product] = 0
                global_total_grains_by_product[found_product] += total_grains

    # ==========================================
    # 4. 處理「粒數總計」列 與 「總粒數 / 總包數」 (重要修正！)
    # ==========================================
    
    row_summary = -1
    for r in range(ws.max_row, 1, -1):
        for c in range(1, 10):
            val = str(ws.cell(row=r, column=c).value).strip()
            if "粒數總計" in val:
                row_summary = r
                break
        if row_summary != -1: break

    # A. 填寫「粒數總計」列
    exclude_list = ["多菁", "普通"]
    if row_summary != -1:
        for col in range(1, ws.max_column + 1):
            prod_name = col_product_map.get(col)
            if prod_name:
                if prod_name not in exclude_list:
                    val = global_total_grains_by_product.get(prod_name, 0)
                    safe_write(ws, row_summary, col, val)
                else:
                    safe_write(ws, row_summary, col, "")

    # B. 填寫「總粒數」與「總包數」 (Smart Jump 邏輯)
    grand_total_grains = sum(global_total_grains_by_product.values())
    
    # 從下往上掃描
    for r in range(ws.max_row, 1, -1):
        for c in range(1, 20): # 掃描前 20 欄
            current_cell = ws.cell(row=r, column=c)
            val = str(current_cell.value).strip()
            
            # 判斷是否為關鍵字
            is_total_grains = "總粒數" in val
            is_total_packs = "總包數" in val
            
            if is_total_grains or is_total_packs:
                # --- 核心修改：判斷要寫入哪一欄 ---
                target_col = c + 1 # 預設是右邊一格
                
                # 檢查當前格子 (標題格) 是否有被合併
                # 如果是 AB 合併，那 c=1 (A)，但我們應該跳過 A, B，寫入 C (3)
                for rng in ws.merged_cells.ranges:
                    if current_cell.coordinate in rng:
                        # 找到了！標題是被合併的
                        # 我們的目標是：合併範圍的最右邊 (max_col) 的下一格
                        target_col = rng.max_col + 1
                        break
                
                # 寫入正確的數值
                if is_total_grains:
                    safe_write(ws, r, target_col, grand_total_grains)
                elif is_total_packs:
                    safe_write(ws, r, target_col, global_total_packs_all)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- 3. Streamlit 介面 (維持不變) ---
st.set_page_config(page_title="檳榔報表生成器 (v11 座標修正版)", layout="wide")
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

st.markdown("### 3. 設定每包粒數")
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
                    st.success("報表生成成功！總粒數與總包數位置已修正。")
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
