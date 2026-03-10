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
        # 確保密碼是字串且去掉前後空格
        office_file.load_key(password=str(password).strip())
        office_file.decrypt(decrypted_buffer)
        decrypted_buffer.seek(0)
        return decrypted_buffer
    except Exception:
        # 解密失敗（密碼不對或本來就沒密碼），回傳原始內容
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
    except Exception:
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
    uploaded_zip = st.file_uploader
