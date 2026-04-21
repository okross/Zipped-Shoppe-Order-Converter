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

EXCLUDE_ITEMS = ["勿拍", "補拍", "補發", "直播下單", "破損鏈接", "破損鏈結", "售後鏈接", "售後鏈結", "直播台", "直播"]

# --- 功能函式 ---

def try_decrypt(file_content, passwords):
    """
    嘗試使用多組密碼解密。
    passwords: list of strings
    """
    # 移除重複值並過濾掉空字串
    pw_list = list(set([str(p).strip() for p in passwords if p]))
    
    # 優先嘗試「不使用密碼」
    try:
        office_file = msoffcrypto.OfficeFile(io.BytesIO(file_content))
        if not office_file.is_encrypted():
            return io.BytesIO(file_content)
        
        # 如果有加密，依序嘗試提供的密碼
        for pw in pw_list:
            try:
                decrypted_buffer = io.BytesIO()
                office_file.load_key(password=pw)
                office_file.decrypt(decrypted_buffer)
                decrypted_buffer.seek(0)
                return decrypted_buffer
            except Exception:
                continue # 密碼錯誤，嘗試下一個
                
    except Exception:
        pass
    
    # 若全部失敗，回傳原始流（交給後續讀取處理，若真有鎖則會報錯）
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
1. **雙重嘗試解密**：自動嘗試資料夾名稱之 **左 6 碼** 與 **右 6 碼** 作為密碼。
2. **自動過濾**：排除退貨、取消及補拍、直播等特殊商品。
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
        
        with st.spinner("正在解析壓縮檔並多重嘗試解密中..."):
            try:
                with zipfile.ZipFile(uploaded_zip) as z:
                    for file_path in z.namelist():
                        if file_path.endswith('.xlsx') and not file_path.split('/')[-1].startswith('._'):
                            
                            path_parts = file_path.split('/')
                            passwords_to_try = []
                            
                            if len(path_parts) > 1:
                                folder_name = path_parts[-2] 
                                # 這裡加入左6碼與右6碼
                                if len(folder_name) >= 6:
                                    passwords_to_try.append(folder_name[:6])   # 左6
                                    passwords_to_try.append(folder_name[-6:])  # 右6
                                else:
                                    passwords_to_try.append(folder_name)       # 長度不足則直接用原名
                            
                            with z.open(file_path) as f:
                                content = f.read()
                                # 呼叫新版的解密函式
                                decrypted_f = try_decrypt(content, passwords_to_try)
                                df_piece = process_excel(decrypted_f)
                                if df_piece is not None:
                                    all_dfs.append(df_piece)
            except Exception as zip_err:
                st.error(f"讀取 ZIP 檔時出錯: {zip_err}")

        if not all_dfs:
            st.error("未找到可讀取的 Excel 檔案，請確認密碼是否正確或檔案是否損毀。")
        else:
            final_df = pd.concat(all_dfs, ignore_index=True)

            # --- 資料清洗邏輯 ---
            if filter_status and "订单状态" in final_df.columns:
                p_status = '取消|退款|退貨|不成立'
                final_df = final_df[~final_df["订单状态"].astype(str).str.contains(p_status, na=False)]

            if "商品名称" in final_df.columns:
                p_items = '|'.join(EXCLUDE_ITEMS)
                final_df = final_df[~final_df["商品名称"].astype(str).str.contains(p_items, na=False)]

            if "快递单号" in final_df.columns:
                final_df = final_df.dropna(subset=["快递单号"])
                final_df = final_df[final_df["快递单号"].astype(str).str.strip() != ""]
            
            if "订单编号" in final_df.columns:
                final_df = final_df.drop_duplicates(subset=["订单编号"], keep='first')

            # 4. 舊訂單排除 (350天)
            excluded_count = 0
            if "订单日期" in final_df.columns:
                final_df["订单日期_dt"] = pd.to_datetime(final_df["订单日期"], errors='coerce')
                cutoff = datetime.now() - timedelta(days=350)
                before_len = len(final_df)
                final_df = final_df[final_df["订单日期_dt"] >= cutoff]
                excluded_count = before_len - len(final_df)

            # --- 金額與單價計算 ---
            final_df['買家總支付金額'] = pd.to_numeric(final_df.get('買家總支付金額', 0), errors='coerce').fillna(0)
            final_df['數量'] = pd.to_numeric(final_df.get('數量', 1), errors='coerce').fillna(1)
            final_df['unit_price'] = (final_df['買家總支付金額'] / final_df['數量'].replace(0, 1)).round(2)

            # --- 建立輸出格式 ---
            result_df = pd.DataFrame()
            result_df['订单编号'] = final_df['订单编号']
            result_df['订单日期'] = final_df['订单日期_dt'].dt.strftime('%Y-%m-%d')
            result_df['订单币种'] = 'TWD'
            result_df['订单金额'] = final_df['買家總支付金額']
            result_df['商品名称'] = final_df.get('商品名称', '')
            result_df['商品数量'] = final_df['數量']
            result_df['商品单价'] = final_df['unit_price']
            result_df['店铺网址'] = shop_url
            result_df['快递单号'] = final_df.get('快递单号', '')
            result_df['物流企业名称'] = final_df.get('物流企业名称', '')
            result_df['电商平台英文名称'] = 'Shopee'

            headers = result_df.columns.tolist()
            ver_row = ["version", "20201013"] + [""] * (len(headers) - 2)
            final_out = [ver_row, headers] + result_df.values.tolist()
            output_df = pd.DataFrame(final_out)

            xlsx_io = io.BytesIO()
            with pd.ExcelWriter(xlsx_io, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False, header=False)
            
            st.success(f"✅ 轉換成功！總筆數：{len(result_df)}，已排除過舊訂單：{excluded_count} 筆。")
            st.download_button(
                label="📥 下載轉換後的 Excel",
                data=xlsx_io.getvalue(),
                file_name=f"Shopee匯出_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
