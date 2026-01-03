import streamlit as st
import pandas as pd
import io
import os
from openpyxl import load_workbook

# --- 1. 核心功能：終極讀取邏輯 (位置鎖定版) ---
def load_and_fix_csv(uploaded_file):
    """
    不管編碼多亂，強制用欄位位置 (Index) 來抓資料。
    """
    bytes_data = uploaded_file.getvalue()
    
    # 準備多種解碼方式來嘗試
    decoding_attempts = []
    
    # 1. 針對你的亂碼特徵 (UTF-8 誤讀 Big5) 的專屬修復
    try:
        text_utf8 = bytes_data.decode('utf-8')
        if '©±' in text_utf8: # 這是你檔案裡 "店名" 的亂碼特徵
             try:
                 # 嘗試還原成中文
                 fixed = text_utf8.encode('latin1').decode('cp950', errors='ignore')
                 decoding_attempts.append(fixed)
             except: pass
        decoding_attempts.append(text_utf8) # 也試試原本的
    except:
        pass
        
    # 2. 傳統中文編碼 (CP950/Big5)
    try:
        decoding_attempts.append(bytes_data.decode('cp950', errors='ignore'))
    except:
        pass
        
    # 3. 英文/原始編碼 (保底，至少不會報錯)
    try:
        decoding_attempts.append(bytes_data.decode('latin1', errors='ignore'))
    except:
        pass

    # 開始逐一測試
    for content in decoding_attempts:
        try:
            # 關鍵修正：
            # 1. sep=',': 強制指定逗號分隔，解決 "Expected 1 fields" 錯誤
            # 2. on_bad_lines='skip': 遇到壞掉的行 (如結尾的加總說明) 直接跳過，不准報錯
            df = pd.read_csv(io.StringIO(content), sep=',', on_bad_lines='skip')
            
            # 檢查欄位數量是否足夠 (你需要抓到第 4 欄)
            if df.shape[1] < 4:
                continue
                
            # --- 欄位鎖定策略 ---
            # 不管標題叫 '©±¦W' 還是 '店名'，我們直接抓位置
            # Index 1 = 店名, Index 2 = 品名, Index 3 = 售量
            
            target_df = df.iloc[:, [1, 2, 3]].copy()
            target_df.columns = ['店名', '品名', '售量'] # 強制改名
            
            # 簡單驗證：售量那一欄應該要有數字
            # 我們試著把售量轉數字，如果成功轉換的比例高，就代表抓對了
            numeric_check = pd.to_numeric(target_df['售量'], errors='coerce')
            if numeric_check.notna().sum() > 0:
                # 清理資料
                target_df['售量'] = numeric_check.fillna(0)
                target_df = target_df.dropna(subset=['店名']) # 店名不能是空的
                return target_df
                
        except Exception:
            continue # 換下一個編碼試試
            
    # 如果試了所有方法都失敗
    st.error(f"檔案 {uploaded_file.name} 徹底讀取失敗，請確認它是否為逗號分隔的 CSV/XLS。")
    return None

# --- 2. 核心功能：處理 Excel 模板 ---
def fill_excel_template(template_path_or_file, combined_df, grains_per_pack_map):
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
        # 累加
        data_dict[store][product] = data_dict[store].get(product, 0) + sales

    # 掃描 header (假設在第 1~10 列之間)
    header_row = 3
    for r in range(1, 10):
        # 找任何看起來像是 "店名" 的格子 (有些模板可能有空白)
        val = ws.cell(row=r, column=1).value
        if val and "店" in str(val): 
            header_row = r
            break
            
    # 掃描品名欄位
    product_col_map = {}
    for col in range(2, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col).value
        if val and isinstance(val, str):
            product_name = val.strip()
            # 排除 "售量" 字眼，剩下的如果是我們的產品名，就記錄下來
            if "售" not in product_name and product_name in grains_per_pack_map:
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
            # A. 綠色字：每包粒數 (填在品名欄)
            grains_setting = grains_per_pack_map.get(product, 0)
            ws.cell(row=row_packs, column=col_idx).value = grains_setting
            
            # B. 紅色字：總銷售包數 (填在售量欄)
            total_packs = total_sales_packs.get(product, 0)
            ws.cell(row=row_packs, column=col_idx + 1).value = total_packs

            # C. 藍色字：總銷售粒數
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

DEFAULT_TEMPLATE = "檳榔銷售統計.xlsx"

col1, col2 = st.columns([1, 2])

with col1:
    st.markdown("### 1. 模板設定")
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
    source_files = st.file_uploader("請一次上傳所有數據檔案", type=["csv", "xls"], accept_multiple_files=True)

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
        st.error("找不到模板檔案！")
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
                    st.error(f"填寫 Excel 時發生錯誤: {e}")
            else:
                st.error("所有檔案讀取失敗，請檢查檔案是否為正確的 CSV/XLS 格式。")
