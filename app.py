import streamlit as st
import pandas as pd
import io
import os
from openpyxl import load_workbook

# --- 1. 核心功能：全能讀取與修復 (支援 Excel 與 CSV) ---
def load_and_fix_smart(uploaded_file):
    """
    自動判斷副檔名：
    - xls/xlsx: 使用 pandas 原生 excel 讀取 (解決二進位亂碼問題)
    - csv: 使用智慧解碼 (解決 Big5/UTF-8 混亂問題)
    最後統一輸出格式標準的 DataFrame
    """
    file_name = uploaded_file.name
    file_ext = os.path.splitext(file_name)[1].lower()
    
    df = None
    msg = ""

    # ==========================
    # 分流 1: 處理 Excel (.xls, .xlsx)
    # ==========================
    if file_ext in ['.xls', '.xlsx']:
        try:
            if file_ext == '.xls':
                # 舊版 Excel (Binary)
                df = pd.read_excel(uploaded_file, engine='xlrd')
            else:
                # 新版 Excel (OpenXML)
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            
            msg = "Excel Read Success"
        except Exception as e:
            return None, f"Excel 讀取失敗 (請確認是否安裝 xlrd): {e}"

    # ==========================
    # 分流 2: 處理 CSV / 文字檔
    # ==========================
    else:
        bytes_data = uploaded_file.getvalue()
        content = ""
        decoded_method = ""
        
        # --- A. 解碼階段 ---
        try:
            text_utf8 = bytes_data.decode('utf-8')
            # 偵測是否為「假 UTF-8 真 Big5」(Mojibake)
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

        # --- B. 標題定位 ---
        lines = content.splitlines()
        header_row_index = -1
        
        for i, line in enumerate(lines[:20]): 
            if "店名" in line and "售量" in line:
                header_row_index = i
                break
                
        if header_row_index == -1:
            return None, f"找不到 CSV 標題列 (使用 {decoded_method})。預覽：{content[:50]}"

        # --- C. 讀取數據 ---
        try:
            valid_content = "\n".join(lines[header_row_index:])
            df = pd.read_csv(io.StringIO(valid_content))
            msg = "CSV Read Success"
        except Exception as e:
            return None, f"解析 CSV 失敗: {e}"

    # ==========================
    # 共同階段: 資料清洗與標準化
    # ==========================
    if df is not None:
        try:
            target_df = pd.DataFrame()

            # 確保欄位名稱存在 (去除前後空白)
            df.columns = [str(c).strip() for c in df.columns]

            # 狀況 1: 欄位名稱正確
            if '店名' in df.columns and '售量' in df.columns:
                target_df = df
            
            # 狀況 2: 欄位名稱跑掉，但欄位數量夠 (Excel 沒讀到 header 時)
            elif df.shape[1] >= 4:
                # 假設順序：店別, 店名, 品名, 售量...
                target_df = df.iloc[:, [1, 2, 3]].copy()
                target_df.columns = ['店名', '品名', '售量']
            else:
                return None, f"欄位識別失敗，請確認檔案包含「店名」與「售量」欄位。"

            # 清理「售量」：轉數字，非數字補 0
            target_df['售量'] = pd.to_numeric(target_df['售量'], errors='coerce').fillna(0)
            
            # 清理「店名」：去除空值
            target_df = target_df.dropna(subset=['店名'])
            
            # 過濾掉可能重複讀到的標題行 (Excel 合併時常見問題)
            target_df = target_df[target_df['店名'].astype(str).str.contains("店名") == False]
            
            return target_df, "Success"
            
        except Exception as e:
            return None, f"資料標準化失敗: {e}"
            
    return None, "Unknown Error"


# --- 2. 核心功能：填寫 Excel (維持不變) ---
def fill_excel_template(template_path_or_file, combined_df, grains_per_pack_map):
    if isinstance(template_path_or_file, str):
        wb = load_workbook(template_path_or_file)
    else:
        wb = load_workbook(template_path_or_file)
    ws = wb.active

    data_dict = {}
    for index, row in combined_df.iterrows():
        store = str(row['店名']).strip()
        product = str(row['品名']).strip()
        sales = row['售量']
        
        if store not in data_dict:
            data_dict[store] = {}
        data_dict[store][product] = data_dict[store].get(product, 0) + sales

    # 定位 Header
    header_row = 3
    for r in range(1, 10):
        val = ws.cell(row=r, column=1).value
        if val and "店" in str(val):
            header_row = r
            break
            
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
st.set_page_config(page_title="檳榔報表生成器 (v7 全能版)", layout="wide")
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
    # 這裡增加了 xlsx 支援
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
        with st.spinner("正在解析數據..."):
            all_data = []
            error_logs = []
            
            for f in source_files:
                # 改用新的函式
                df, msg = load_and_fix_smart(f)
                if df is not None:
                    all_data.append(df)
                else:
                    error_logs.append(f"❌ {f.name}: {msg}")
            
            # 顯示錯誤日誌 (如果有)
            if error_logs:
                with st.expander("⚠️ 部分檔案讀取失敗 (點擊查看詳情)"):
                    for log in error_logs:
                        st.code(log)
            
            if all_data:
                combined_df = pd.concat(all_data, ignore_index=True)
                st.info(f"✅ 成功讀取 {len(combined_df)} 筆資料。")
                
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
                st.error("沒有任何檔案被成功讀取。請查看上方的錯誤日誌。")
