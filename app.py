import streamlit as st
import pandas as pd
import io
import os
from openpyxl import load_workbook

# --- 1. 核心功能：亂碼修復 + 暴力過濾雜訊 ---
def load_and_fix_csv_robust(uploaded_file):
    """
    1. 修復編碼 (UTF-8 -> Latin1 -> CP950)
    2. 過濾掉沒有逗號的雜訊行 (解決 Expected 1 fields 錯誤)
    """
    file_name = uploaded_file.name
    bytes_data = uploaded_file.getvalue()
    
    # --- 步驟 A: 解碼 (找回中文字) ---
    content = ""
    # 策略 1: 針對你的檔案特徵 (UTF-8 亂碼還原)
    try:
        text_utf8 = bytes_data.decode('utf-8')
        if '©±' in text_utf8: # 偵測到你的亂碼特徵
            # 這就是你要的 "對照表" 邏輯：反向編碼回 Latin1，再用 CP950 解開
            content = text_utf8.encode('latin1').decode('cp950', errors='ignore')
        else:
            content = text_utf8
    except:
        # 策略 2: 如果上面失敗，直接試試 CP950
        try:
            content = bytes_data.decode('cp950', errors='ignore')
        except:
            content = bytes_data.decode('latin1', errors='ignore')

    # --- 步驟 B: 清洗數據 (解決格式錯誤) ---
    # 這是這次修正的關鍵：不要直接丟給 Pandas 讀，我們先把壞掉的行踢掉
    valid_lines = []
    lines = content.splitlines()
    
    for line in lines:
        # 簡單判斷：有效的資料行至少要有 2 個以上的逗號 (店名, 品名, 售量...)
        if line.count(',') >= 2:
            valid_lines.append(line)
            
    if not valid_lines:
        st.error(f"檔案 {file_name} 內容看起來是空的或格式全錯。")
        return None

    # 重組回 CSV 字串
    clean_content = "\n".join(valid_lines)

    # --- 步驟 C: 轉成表格 ---
    try:
        # 這次我們自己指定欄位名稱，不管它標題寫什麼亂碼，反正順序是固定的
        # header=0 表示第一行是標題 (我們會把它覆蓋掉)
        df = pd.read_csv(io.StringIO(clean_content), header=0)
        
        # 你的檔案結構：第2欄=店名, 第3欄=品名, 第4欄=售量 (Python index 從 0 開始，所以是 1, 2, 3)
        # 先檢查欄位數夠不夠
        if df.shape[1] < 4:
            # 有時候標題行被過濾掉了，試試看有沒有可能是無標題狀態
            df = pd.read_csv(io.StringIO(clean_content), header=None)
        
        if df.shape[1] >= 4:
            # 強制鎖定我們要的欄位
            target_df = df.iloc[:, [1, 2, 3]].copy()
            target_df.columns = ['店名', '品名', '售量']
            
            # 清理：確保售量是數字
            target_df['售量'] = pd.to_numeric(target_df['售量'], errors='coerce').fillna(0)
            target_df = target_df.dropna(subset=['店名']) # 去除店名空的行
            
            # 排除標題行本身被當成資料讀進來的情況 (如果店名那欄寫著 "店名")
            target_df = target_df[target_df['店名'].astype(str).str.contains("店名|©±") == False]
            
            return target_df
        else:
            st.warning(f"檔案 {file_name} 欄位不足，無法解析。")
            return None

    except Exception as e:
        st.error(f"檔案 {file_name} 解析失敗: {e}")
        return None

# --- 2. 核心功能：填寫 Excel (維持不變) ---
def fill_excel_template(template_path_or_file, combined_df, grains_per_pack_map):
    if isinstance(template_path_or_file, str):
        wb = load_workbook(template_path_or_file)
    else:
        wb = load_workbook(template_path_or_file)
    ws = wb.active

    # 1. 數據匯總
    data_dict = {}
    for index, row in combined_df.iterrows():
        store = str(row['店名']).strip()
        product = str(row['品名']).strip()
        sales = row['售量']
        
        if store not in data_dict:
            data_dict[store] = {}
        data_dict[store][product] = data_dict[store].get(product, 0) + sales

    # 2. 定位標題
    header_row = 3
    for r in range(1, 10):
        val = ws.cell(row=r, column=1).value
        if val and "店" in str(val):
            header_row = r
            break
            
    # 3. 定位品名欄
    product_col_map = {}
    for col in range(2, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col).value
        if val and isinstance(val, str):
            product_name = val.strip()
            if "售" not in product_name and product_name in grains_per_pack_map:
                product_col_map[product_name] = col

    total_sales_packs = {p: 0 for p in product_col_map}
    row_packs = None
    row_grains = None
    
    # 4. 填寫內容
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

    # 5. 填寫統計
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
st.set_page_config(page_title="檳榔報表生成器 (強力版)", layout="wide")
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
    source_files = st.file_uploader("上傳所有數據檔案 (特幼, 雙子星...)", type=["csv", "xls"], accept_multiple_files=True)

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
    if not template_file and not os.path.exists(DEFAULT_TEMPLATE):
        st.error("找不到模板檔案！")
    elif not source_files:
        st.error("請上傳原始數據檔案。")
    else:
        # 如果使用者沒上傳新模板，且有勾選預設，則使用預設路徑
        current_template = template_file if template_file else DEFAULT_TEMPLATE
        
        with st.spinner("正在強力解析數據..."):
            all_data = []
            for f in source_files:
                df = load_and_fix_csv_robust(f)
                if df is not None:
                    all_data.append(df)
            
            if all_data:
                combined_df = pd.concat(all_data, ignore_index=True)
                
                # 顯示一下讀取到的數據量，讓你知道有沒有成功
                st.info(f"成功讀取 {len(combined_df)} 筆銷售紀錄，正在填寫報表...")
                
                try:
                    result_excel = fill_excel_template(current_template, combined_df, user_grains_setting)
                    st.success("報表生成成功！")
                    st.download_button(
                        label="📥 下載報表",
                        data=result_excel,
                        file_name="已填寫_檳榔銷售統計.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"填寫 Excel 時發生錯誤: {e}")
            else:
                st.error("所有檔案都無法讀取，請確認檔案內容是否正確。")
