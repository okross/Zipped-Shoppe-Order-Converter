import streamlit as st
import pandas as pd
import io
import zipfile
import msoffcrypto
from datetime import datetime, timedelta

# --- 配置區 ---
COLUMN_MAPPING = {
    "订单编号": ["訂單編號", "Order ID"],
    "订单日期": ["訂單日期", "訂單成立日期", "下單時間"],
    "商品名称": ["商品名稱", "商品項目"],
    "快递单号": ["包裹號碼", "包裹查詢號碼", "寄件單號"],
    "物流企业名称": ["寄送方式", "運送方式"],
    "订单状态": ["訂單狀態", "Order Status"],
    "買家總支付金額": ["買家總支付金額", "買家總支付", "商品總價"],
    "數量": ["數量", "商品數量"]
}

# 商品排除關鍵字
EXCLUDE_ITEMS = ["勿拍", "補拍", "補發", "直播下單", "破損鏈接", "破損鏈結", "售後鏈接", "售後鏈結"]

# --- 功能函式 ---

def try_decrypt(file_content, password):
    """嘗試使用指定密碼解密 Excel，若失敗則回傳原始流"""
    decrypted_buffer = io.BytesIO()
    try:
        office_file = msoffcrypto.OfficeFile(io.BytesIO(file_content))
        # 確保密碼是字串且去掉空格
        office_file.load_key(password=str(password).strip())
        office_file.decrypt(decrypted_buffer)
        decrypted_buffer.seek(0)
        return decrypted_buffer
    except:
        return io.BytesIO(file_content)

def process_excel(file_stream):
    """讀取並轉換單個 Excel 的欄位"""
    try:
        df = pd.read_excel(file_stream, engine='openpyxl')
        df.columns = [str(col).strip().replace('\n', '') for col in df.columns]
        
        found_mapping = {}
        for target, aliases in COLUMN_MAPPING.items():
            for alias in aliases:
                if alias in df.columns:
                    found_mapping[alias] = target
                    break
        
        if found_mapping:
            return df[list(found_mapping.keys())].rename(columns=found_mapping)
    except:
        return None
    return None

# --- Streamlit 網頁介面 ---

st.set_page_config(page_title="Shopee Order Converter (ZIP)", layout="centered")

st.title("📦 Shopee 訂單 ZIP 自動轉換器")
st.markdown("""
本系統會自動讀取 ZIP 內的 Excel 檔案：
1. **自動解密**：取資料夾名稱之**前 6 碼**作為密碼。
2. **自動過濾**：排除退貨、取消及補拍、補發等特殊商品。
""")

with st.form("main_form"):
    shop_url = st.text_input("1. 請輸入店鋪網址 (必填)", placeholder="https://shopee.tw/yourshop")
    filter_status = st.checkbox("2. 自動排除退貨/取消訂單", value=True)
    uploaded_zip = st.file_uploader("3. 上傳 ZIP 壓縮檔", type=["zip"])
    submit = st.form_submit_button("執行轉換")

if submit:
    if not shop_url:
        st.error("請填寫店鋪網址！")
    elif not uploaded_zip:
        st.error("請上傳 ZIP 檔案！")
    else:
        all_dfs = []
        
        with st.spinner("正在處理壓縮檔內的所有資料夾..."):
            with zipfile.ZipFile(uploaded_zip) as z:
                for file_path in z.namelist():
                    # 排除系統垃圾檔案與資料夾路徑本身
                    if file_path.endswith('.xlsx') and not file_path.startswith('__MACOSX'):
                        
                        # 邏輯：從路徑取得資料夾名稱並抓取前六碼
                        path_parts = file_path.split('/')
                        password = ""
                        if len(path_parts) > 1:
                            folder_name = path_parts[-2] # 取得檔案所在的資料夾名
                            password = folder_name[:6]
                        
                        # 讀取並嘗試解密
                        with z.open(file_path) as f:
                            content = f.read()
                            decrypted_f = try_decrypt(content, password)
                            df_piece = process_excel(decrypted_f)
                            if df_piece is not None:
                                all_dfs.append(df_piece)

        if not all_dfs:
            st.error("未找到可匹配的 Excel 檔案，請確認檔案內容或欄位名稱。")
        else:
            final_df = pd.concat(all_dfs, ignore_index=True)

            # --- 資料清洗邏輯 ---
            # 1. 狀態過濾
            if filter_status and "订单状态" in final_df.columns:
                p_status = '|'.join(['取消', '退款', '退貨', '不成立'])
                final_df = final_df[~final_df["订单状态"].astype(str).str.contains(p_status, na=False)]

            # 2. 商品名稱過濾 (包含你指定的所有關鍵字)
            if "商品名称" in final_df.columns:
                p_items = '|'.join(EXCLUDE_ITEMS)
                final_df = final_df[~final_df["商品名称"].astype(str).str.contains(