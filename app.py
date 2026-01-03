import streamlit as st
import pandas as pd
import io
import os
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

# --- 1. 核心功能：全能讀取與修復 ---
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

# --- 2. 核心功能：填寫 Excel (V15 順序暴力填充版) ---
def fill_excel_template_sequential(template_path_or_file, combined_df, grains_per_pack_map):
    if isinstance(template_path_or_file, str):
        wb = load_workbook(template_path_or_file)
    else:
        wb = load_workbook(template_path_or_file)
    ws = wb.active
    
    update_log = [] 

    # ==========================================
    # 步驟 1: 整理來源數據 (按產品分組，保持順序)
    # ==========================================
    # 結構： { '特幼': [100, 26, 66, ...], '多粒': [25, 32, ...] }
    sales_lists_by_product = {}
    
    # 這裡假設 combined_df 的順序就是 user 上傳的順序 (或是 Excel 裡的原始順序)
    # 為了保險，我們針對每一個 product 建立一個列表
    
    # 取得所有出現過的產品
    unique_products = combined_df['品名'].unique()
    
    for prod in unique_products:
        prod_key = str(prod).strip()
        # 找出該產品的所有銷售數據 (依原始順序)
        sales_series = combined_df[combined_df['品名'] == prod]['售量'].tolist()
        
        # 進行模糊匹配，對應到 grains_per_pack_map 的 key
        matched_key = prod_key
        for key in grains_per_pack_map.keys():
            if key in prod_key:
                matched_key = key
                break
        
        if matched_key not in sales_lists_by_product:
            sales_lists_by_product[matched_key] = []
        
        sales_lists_by_product[matched_key].extend(sales_series)

    # ==========================================
    # 步驟 2: 定位 Excel 結構
    # ==========================================
    header_row = 3
    store_col_index = 1 
    
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
    
    # 找出所有 (品名欄, 售量欄)
    col_pairs = [] 
    for c in range(1, ws.max_column + 1):
        val1 = str(ws.cell(row=header_row, column=c).value).strip()
        val2 = str(ws.cell(row=header_row, column=c+1).value).strip()
        if "品名" in val1 and "售量" in val2:
            col_pairs.append((c, c+1))

    # ==========================================
    # 步驟 3: 初始化 (清空舊數據) - Robust 關鍵
    # ==========================================
    # 我們把所有欄位的售量都清空，避免沒有讀到的產品殘留舊值
    for r in range(header_row + 1, ws.max_row + 1):
        cell_store = ws.cell(row=r, column=store_col_index).value
        if not cell_store or "銷售" in str(cell_store) or "合計" in str(cell_store): continue
        
        for (prod_col, sales_col) in col_pairs:
             safe_write(ws, r, sales_col, 0) # 先全部歸零

    # ==========================================
    # 步驟 4: 暴力依序填充 (Sequential Paste)
    # ==========================================
    # 我們需要知道 Excel 裡的每一個 Column 是屬於哪個產品
    # 這裡採用動態偵測：掃描每一行，看該產品欄位的品名是什麼，然後從清單中拿出下一個數字填入
    
    # 為了處理「多菁/普通」這種只有部分店有的情況：
    # 假設 Excel 裡這欄的格子是空的或是特定標記？
    # 不，通常 Excel 模板每個店都有格子。
    # 如果使用者說 "照順序貼"，代表來源資料的筆數 = Excel 裡的店家數 (或者對應的店家數)
    # 我們維護一個 index 指標： { '特幼': 0, '多粒': 0 ... } 指向目前填到第幾個數字
    
    current_idx_map = {k: 0 for k in sales_lists_by_product.keys()}
    
    for r in range(header_row + 1, ws.max_row + 1):
        cell_store = ws.cell(row=r, column=store_col_index).value
        
        if not cell_store: continue
        if "銷售" in str(cell_store) or "合計" in str(cell_store): continue
        
        # 對這一列的每一組 (Prod, Sales)
        for (prod_col, sales_col) in col_pairs:
            cell_prod = ws.cell(row=r, column=prod_col).value
            if not cell_prod: continue
            prod_name_in_excel = str(cell_prod).strip()
            
            # 辨識這是哪個產品
            target_key = None
            for key in grains_per_pack_map.keys():
                if key in prod_name_in_excel:
                    target_key = key
                    break
            
            # 如果我們手上有這個產品的數據清單
            if target_key and target_key in sales_lists_by_product:
                data_list = sales_lists_by_product[target_key]
                idx = current_idx_map[target_key]
                
                # 還有彈藥嗎？
                if idx < len(data_list):
                    val_to_write = data_list[idx]
                    safe_write(ws, r, sales_col, val_to_write)
                    current_idx_map[target_key] += 1 # 準備填下一個
                else:
                    # 彈藥用盡 (可能來源資料比 Excel 店家少)，保持 0
                    pass

    # ==========================================
    # 步驟 5: 統計與結算 (同前版)
    # ==========================================
    global_total_grains_by_product = {} 
    global_total_packs_all = 0

    pack_rows = []
    for r in range(1, ws.max_row + 1):
        val = ws.cell(row=r, column=store_col_index).value
        if val and "銷售包數" in str(val):
            pack_rows.append(r)

    for r_pack in pack_rows:
        r_grain = -1
        next_cell = ws.cell(row=r_pack + 1, column=store_col_index).value
        if next_cell and "銷售粒數" in str(next_cell):
            r_grain = r_pack + 1

        for (prod_col, sales_col) in col_pairs:
            found_product = None
            for offset in range(1, 6):
                val = ws.cell(row=r_pack - offset, column=prod_col).value
                if val and isinstance(val, str) and len(val) > 1:
                    for key in grains_per_pack_map.keys():
                        if key in val:
                            found_product = key
                            break
                    if found_product: break
            
            if found_product:
                setting_val = grains_per_pack_map.get(found_product)
                safe_write(ws, r_pack, prod_col, setting_val)
                
                # 重新計算紅色總和 (因為我們剛剛填入了數據)
                current_red_sum = 0
                for offset in range(1, 20):
                    r_scan = r_pack - offset
                    if r_scan <= header_row: break
                    val = ws.cell(row=r_scan, column=sales_col).value
                    if isinstance(val, (int, float)):
                        current_red_sum += val
                
                safe_write(ws, r_pack, sales_col, current_red_sum)
                global_total_packs_all += current_red_sum
                
                total_grains = current_red_sum * setting_val
                if r_grain != -1:
                    safe_write(ws, r_grain, sales_col, total_grains)
                
                if found_product not in global_total_grains_by_product:
                    global_total_grains_by_product[found_product] = 0
                global_total_grains_by_product[found_product] += total_grains

    # 步驟 6: 總結算
    row_summary = -1
    for r in range(ws.max_row, 1, -1):
        for c in range(1, 10):
            val = str(ws.cell(row=r, column=c).value).strip()
            if "粒數總計" in val:
                row_summary = r
                break
        if row_summary != -1: break

    exclude_list = ["多菁", "普通"]

    if row_summary != -1:
        for (prod_col, sales_col) in col_pairs:
            target_product = None
            if pack_rows:
                first_pack_row = pack_rows[0]
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
                safe_write(ws, row_summary, sales_col, "")

    # B. 總粒數與總包數
    grand_total_grains = sum(global_total_grains_by_product.values())
    
    for r in range(ws.max_row, 1, -1):
        for c in range(1, 20): 
            current_cell = ws.cell(row=r, column=c)
            val = str(current_cell.value).strip()
            
            is_total_grains = "總粒數" in val
            is_total_packs = "總包數" in val
            
            if is_total_grains or is_total_packs:
                target_col = c + 1
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
    return output, update_log

# --- 3. Streamlit 介面 ---
st.set_page_config(page_title="檳榔報表生成器 (v15 順序暴力填充版)", layout="wide")
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
                
                with st.expander("🔍 查看讀取數據詳情 (確認順序是否正確)"):
                    st.dataframe(combined_df)

                try:
                    result_excel, logs = fill_excel_template_sequential(current_template, combined_df, user_grains_setting)
                    st.success("報表生成成功！已使用順序強制填充模式。")
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
