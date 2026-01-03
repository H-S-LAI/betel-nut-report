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

# --- 功能：安全寫入 ---
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

# --- 2. 核心功能：填寫 Excel (V13 寬版掃描架構) ---
def fill_excel_template(template_path_or_file, combined_df, grains_per_pack_map):
    if isinstance(template_path_or_file, str):
        wb = load_workbook(template_path_or_file)
    else:
        wb = load_workbook(template_path_or_file)
    ws = wb.active

    # ==========================================
    # 準備數據字典
    # ==========================================
    global_total_grains_by_product = {} 
    global_total_packs_all = 0
    
    data_dict = {}
    for index, row in combined_df.iterrows():
        store = str(row['店名']).strip()
        product = str(row['品名']).strip()
        sales = row['售量']
        
        if store not in data_dict:
            data_dict[store] = {}
        
        # 模糊匹配
        matched_key = product
        for key in grains_per_pack_map.keys():
            if key in product:
                matched_key = key
                break
        data_dict[store][matched_key] = data_dict[store].get(matched_key, 0) + sales

    # ==========================================
    # 步驟 1: 定位與掃描欄位結構 (Wide Scan)
    # ==========================================
    header_row = 3
    store_col_index = 1 # 預設 A 欄
    
    # 找 Header Row
    for r in range(1, 10):
        found = False
        for c in range(1, 10):
            val = ws.cell(row=r, column=c).value
            if val and "店" in str(val):
                header_row = r
                store_col_index = c
                found = True
                break
        if found: break
    
    # 找出所有的 (品名欄, 售量欄) 配對
    # 邏輯：只要標題是 "品名" 且右邊是 "售量"，就是一組
    col_pairs = [] # list of tuple (prod_col, sales_col)
    
    for c in range(1, ws.max_column + 1):
        val1 = str(ws.cell(row=header_row, column=c).value).strip()
        val2 = str(ws.cell(row=header_row, column=c+1).value).strip()
        
        if "品名" in val1 and "售量" in val2:
            col_pairs.append((c, c+1))

    # ==========================================
    # 步驟 2: 填寫銷售數據 (Main Data Filling)
    # ==========================================
    # 從 Header 下一行開始掃描每一列
    for r in range(header_row + 1, ws.max_row + 1):
        cell_store = ws.cell(row=r, column=store_col_index).value
        
        # 跳過空行或特殊行
        if not cell_store: continue
        if "銷售" in str(cell_store) or "合計" in str(cell_store): continue
        
        store_name = str(cell_store).strip()
        
        # 對這一列的每一組 (Prod, Sales) 進行檢查與填寫
        for (prod_col, sales_col) in col_pairs:
            # 讀取 Excel 預設的品名
            cell_prod = ws.cell(row=r, column=prod_col).value
            if not cell_prod: continue
            
            prod_name_in_excel = str(cell_prod).strip()
            
            # 在數據庫中尋找
            if store_name in data_dict:
                sales_val = 0
                # 嘗試匹配
                for key_prod in data_dict[store_name]:
                    if key_prod in prod_name_in_excel or prod_name_in_excel in key_prod:
                        sales_val = data_dict[store_name][key_prod]
                        break
                
                # 如果有數據，強制寫入 (這解決了 24 變 100 的問題)
                if sales_val > 0:
                    safe_write(ws, r, sales_col, sales_val)

    # ==========================================
    # 步驟 3: 處理「紅色包數」與「藍色粒數」 (Per-Store Totals)
    # ==========================================
    # 找出所有 "銷售包數" 的行
    pack_rows = []
    for r in range(1, ws.max_row + 1):
        val = ws.cell(row=r, column=store_col_index).value
        if val and "銷售包數" in str(val):
            pack_rows.append(r)

    for r_pack in pack_rows:
        r_grain = -1
        # 找下一行的 "銷售粒數"
        next_cell = ws.cell(row=r_pack + 1, column=store_col_index).value
        if next_cell and "銷售粒數" in str(next_cell):
            r_grain = r_pack + 1

        # 針對每一組 (Prod, Sales) 進行統計
        for (prod_col, sales_col) in col_pairs:
            # 往上找品名 (最多找 5 格)
            found_product = None
            for offset in range(1, 6):
                val = ws.cell(row=r_pack - offset, column=prod_col).value
                if val and isinstance(val, str) and len(val) > 1:
                    # 檢查是不是已知的產品
                    for key in grains_per_pack_map.keys():
                        if key in val:
                            found_product = key
                            break
                    if found_product: break
            
            if found_product:
                # 1. 寫入綠色 (粒數設定) -> 寫在 Product Column
                setting_val = grains_per_pack_map.get(found_product)
                safe_write(ws, r_pack, prod_col, setting_val)
                
                # 2. 計算紅色 (該區塊總包數) -> 讀取 Sales Column
                current_red_sum = 0
                for offset in range(1, 20): # 往上掃描
                    r_scan = r_pack - offset
                    if r_scan <= header_row: break
                    
                    # 確保是同一家店的範圍 (簡單判斷: 左邊有店名或者是數據區)
                    # 這裡直接讀取數值累加
                    val = ws.cell(row=r_scan, column=sales_col).value
                    if isinstance(val, (int, float)):
                        current_red_sum += val
                
                # 寫入紅色 -> 寫在 Sales Column
                safe_write(ws, r_pack, sales_col, current_red_sum)
                
                # 全域累加
                global_total_packs_all += current_red_sum
                
                # 3. 計算與寫入藍色 (總粒數) -> 寫在 Sales Column
                total_grains = current_red_sum * setting_val
                if r_grain != -1:
                    safe_write(ws, r_grain, sales_col, total_grains)
                
                # 全域累加
                if found_product not in global_total_grains_by_product:
                    global_total_grains_by_product[found_product] = 0
                global_total_grains_by_product[found_product] += total_grains

    # ==========================================
    # 步驟 4: 處理最下方的總結算 (Summary Row)
    # ==========================================
    row_summary = -1
    for r in range(ws.max_row, 1, -1):
        for c in range(1, 10):
            val = str(ws.cell(row=r, column=c).value).strip()
            if "粒數總計" in val:
                row_summary = r
                break
        if row_summary != -1: break

    exclude_list = ["多菁", "普通"]

    # A. 填寫粒數總計
    if row_summary != -1:
        # 只填寫 Sales Column，跳過 Product Column
        for (prod_col, sales_col) in col_pairs:
            # 我們需要知道這一欄對應什麼產品？往上找 header 附近的數據
            # 簡單一點：遍歷 global_total_grains_by_product 來匹配
            # 更好的方法：重新確認這一欄的品名
            
            # 往上找 header row 的品名 (雖然 header 寫 "品名"，但我們需要知道是哪個產品)
            # 我們可以用剛剛的邏輯：該 sales_col 對應到的 found_product
            # 為了效率，我們這裡不做複雜回溯，而是假設 column order 沒變
            # 讓我們用一個更穩的方法：遍歷所有 sales_col，往上找產品名
            
            target_product = None
            # 找這一欄上面的紅色格子附近的產品名
            # 找第一個 pack_row
            if pack_rows:
                first_pack_row = pack_rows[0]
                # 往上找
                for offset in range(1, 6):
                    val = ws.cell(row=first_pack_row - offset, column=prod_col).value
                    if val:
                         for key in grains_per_pack_map.keys():
                            if key in str(val):
                                target_product = key
                                break
                    if target_product: break
            
            if target_product and target_product not in exclude_list:
                val = global_total_grains_by_product.get(target_product, 0)
                safe_write(ws, row_summary, sales_col, val)
            else:
                # 不填寫，或填空
                safe_write(ws, row_summary, sales_col, "")

    # B. 填寫總粒數與總包數 (Smart Jump)
    grand_total_grains = sum(global_total_grains_by_product.values())
    
    for r in range(ws.max_row, 1, -1):
        for c in range(1, 20): 
            current_cell = ws.cell(row=r, column=c)
            val = str(current_cell.value).strip()
            
            is_total_grains = "總粒數" in val
            is_total_packs = "總包數" in val
            
            if is_total_grains or is_total_packs:
                target_col = c + 1
                # 檢查合併儲存格跳躍
                for rng in ws.merged_cells.ranges:
                    if current_cell.coordinate in rng:
                        target_col = rng.max_col + 1
                        break
                
                if is_total_grains:
                    safe_write(ws, r, target_col, grand_total_grains)
                elif is_total_packs:
                    safe_write(ws, r, target_col, global_total_packs_all)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- 3. Streamlit 介面 ---
st.set_page_config(page_title="檳榔報表生成器 (v13 寬版掃描版)", layout="wide")
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
                    df['來源檔案'] = f.name 
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
                
                with st.expander("🔍 點擊查看數據詳情"):
                    st.dataframe(combined_df)

                try:
                    result_excel = fill_excel_template(current_template, combined_df, user_grains_setting)
                    st.success("報表生成成功！已修正寬表格讀取問題。")
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
