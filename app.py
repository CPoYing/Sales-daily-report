import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
st.title("Excel 銷售日報表 Mapping 工具")
# 用戶上傳檔案，銷貨退回非必填
sales_file = st.file_uploader("41110000 銷貨收入", type="xlsx")
returns_file = st.file_uploader("41700000 銷貨退回 (選填)", type="xlsx")
zsdc_file = st.file_uploader("zsdc", type="xlsx")
contract_file = st.file_uploader("合約管理", type="xlsx")
product_file = st.file_uploader("產品群", type="xlsx")
quote_file = st.file_uploader("報價資訊 (選填)", type="xlsx")
# 讓使用者設定月份（預設系統當月）
current_month = datetime.datetime.now().month
month = st.number_input("設定月份 (預設當月)", min_value=1, max_value=12, value=current_month)
# 新增兩個輸入數值欄位，改成小數兩位
m1_copper_price = st.number_input("M-1銅價", value=0.00, step=0.01, format="%0.2f")
# 動態多組 M-2 銅價輸入
st.write("M-2銅價組 (可新增多組)")
if 'm2_groups' not in st.session_state:
    st.session_state.m2_groups = [{'month': current_month - 2 if current_month > 2 else 10, 'price': 0.00}]
if st.button("新增 M-2 組"):
    st.session_state.m2_groups.append({'month': current_month - 2 if current_month > 2 else 10, 'price': 0.00})
m2_dict = {}
existing_months = set()
for i, group in enumerate(st.session_state.m2_groups):
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        new_month = st.number_input(f"M-2月份 {i+1}", min_value=1, max_value=12, value=group['month'], key=f"m2_month_{i}")
        if new_month in existing_months:
            st.warning(f"月份 {new_month} 已存在，請選擇其他月份")
        else:
            group['month'] = new_month
            existing_months.add(new_month)
    with col2:
        group['price'] = st.number_input(f"M-2銅價 {i+1}", value=group['price'], step=0.01, format="%0.2f", key=f"m2_price_{i}")
    with col3:
        if st.button("刪除", key=f"del_m2_{i}"):
            del st.session_state.m2_groups[i]
            st.rerun()
    m2_dict[group['month']] = group['price']
# 新增時間區間篩選
start_date = st.date_input("開始日期 (篩選輸入日期)", value=datetime.datetime.now().date() - datetime.timedelta(days=7))
end_date = st.date_input("結束日期 (篩選輸入日期)", value=datetime.datetime.now().date())
if st.button("處理檔案") and sales_file and zsdc_file:
    try:
        # 讀取檔案
        df_sales = pd.read_excel(sales_file, sheet_name='工作表1')
        df_zsdc = pd.read_excel(zsdc_file, sheet_name='工作表1')
        # 如果有退回檔案才讀取，否則空DataFrame
        if returns_file:
            df_returns = pd.read_excel(returns_file, sheet_name='工作表1')
        else:
            df_returns = pd.DataFrame()
        # 如果有合約檔案，讀取
        if contract_file:
            df_contract = pd.read_excel(contract_file, sheet_name='工作表1')
        else:
            df_contract = pd.DataFrame()
        # 如果有產品群檔案，讀取
        if product_file:
            df_product = pd.read_excel(product_file, sheet_name='工作表1')
        else:
            df_product = pd.DataFrame()
        # 如果有報價資訊檔案，讀取
        if quote_file:
            df_quote = pd.read_excel(quote_file, sheet_name='工作表1')
        else:
            df_quote = pd.DataFrame()
        # 篩選銷貨收入和銷貨退回的輸入日期 (L欄: 輸入日期)
        df_sales['輸入日期'] = pd.to_datetime(df_sales['輸入日期'], errors='coerce')
        df_sales = df_sales[(df_sales['輸入日期'] >= pd.to_datetime(start_date)) & (df_sales['輸入日期'] <= pd.to_datetime(end_date))]
        if not df_returns.empty:
            df_returns['輸入日期'] = pd.to_datetime(df_returns['輸入日期'], errors='coerce')
            df_returns = df_returns[(df_returns['輸入日期'] >= pd.to_datetime(start_date)) & (df_returns['輸入日期'] <= pd.to_datetime(end_date))]
        # 重命名欄位
        df_sales = df_sales.rename(columns={'項目.1': 'sales_item'})
        if not df_returns.empty:
            df_returns = df_returns.rename(columns={'項目.1': 'sales_item'})
        # 共同欄位
        common_cols = ['參考文件號碼', '項目', '物料', '工廠', '客戶', '銷售文件', 'sales_item', '以 PCLC 計', '數量', '過帳日期', 'BUn', '輸入日期']
        df_sales_common = df_sales[common_cols]
        df_sales_common['以 PCLC 計'] = -df_sales_common['以 PCLC 計'] # *-1
        df_sales_common['數量'] = -df_sales_common['數量'] # *-1
        if not df_returns.empty:
            df_returns_common = df_returns[common_cols]
            df_returns_common['以 PCLC 計'] = -df_returns_common['以 PCLC 計'] # *-1
            df_returns_common['數量'] = -df_returns_common['數量'] # *-1
        else:
            df_returns_common = pd.DataFrame(columns=common_cols)
        # 合併
        df_combined = pd.concat([df_sales_common, df_returns_common], ignore_index=True)
        # 轉型銷售文件為int，確保匹配
        df_combined['銷售文件'] = pd.to_numeric(df_combined['銷售文件'], errors='coerce').fillna(0).astype(int)
        df_zsdc['先前文件'] = pd.to_numeric(df_zsdc['先前文件'], errors='coerce').fillna(0).astype(int)
        # 建立 key (強制轉整數排除小數點)
        df_combined['參考文件號碼'] = pd.to_numeric(df_combined['參考文件號碼'], errors='coerce').fillna(0).astype(int)
        df_combined['項目'] = pd.to_numeric(df_combined['項目'], errors='coerce').fillna(0).astype(int)
        df_combined['key'] = df_combined['參考文件號碼'].astype(str) + '_' + df_combined['項目'].astype(str)
        df_zsdc['文件'] = pd.to_numeric(df_zsdc['文件'], errors='coerce').fillna(0).astype(int)
        df_zsdc['項目'] = pd.to_numeric(df_zsdc['項目'], errors='coerce').fillna(0).astype(int)
        df_zsdc['key'] = df_zsdc['文件'].astype(str) + '_' + df_zsdc['項目'].astype(str)
        # 處理zsdc重複key（取第一個）
        df_zsdc = df_zsdc.drop_duplicates(subset='key', keep='first')
        # mapping zsdc 欄位（只取品名和單位用銅，物料保留銷貨收入/退回原始值）
        mapping_dict = df_zsdc.set_index('key')[['物料說明', '淨重']].to_dict(orient='index')
        df_combined['品名'] = df_combined['key'].apply(lambda k: mapping_dict.get(k, {}).get('物料說明', ''))
        df_combined['單位用銅'] = df_combined['key'].apply(lambda k: mapping_dict.get(k, {}).get('淨重', ''))
        # 客戶名稱 mapping 用銷售文件 mapping zsdc 先前文件
        sales_doc_dict = df_zsdc.set_index('先前文件')['bill-to-name'].to_dict()
        df_combined['客戶名稱'] = df_combined['銷售文件'].map(sales_doc_dict).fillna('')
        # 合約號碼和採購單 mapping 用相同邏輯
        contract_no_dict = df_zsdc.set_index('先前文件')['合約號碼'].to_dict()
        purchase_order_dict = df_zsdc.set_index('先前文件')['採購單號碼'].to_dict()
        df_combined['合約號碼'] = df_combined['銷售文件'].map(contract_no_dict).fillna('')
        df_combined['採購單'] = df_combined['銷售文件'].map(purchase_order_dict).fillna('')
        # 計算銅量 = 單位用銅 * 數量
        df_combined['單位用銅'] = pd.to_numeric(df_combined['單位用銅'], errors='coerce').fillna(0)
        df_combined['銅量'] = df_combined['單位用銅'] * df_combined['數量']
        # 如果有合約檔案，mapping 指定欄位 (用合約號碼作為key)
        if not df_contract.empty:
            # 轉型合約號碼確保匹配
            df_combined['合約號碼'] = df_combined['合約號碼'].astype(str)
            df_contract['合約編號'] = df_contract['合約編號'].astype(str)
            # 處理合約重複key
            df_contract = df_contract.drop_duplicates(subset='合約編號', keep='first')
            contract_mapping = df_contract.set_index('合約編號')[['產品部', '通路', '部門', '報價單號', '匯率', '業務', '報價銅價']].to_dict(orient='index')
            # 產品部 mapping 到 線種，產品群留空
            df_combined['線種'] = df_combined['合約號碼'].apply(lambda k: contract_mapping.get(k, {}).get('產品部', ''))
            df_combined['產品群'] = ''
            df_combined['通路'] = df_combined['合約號碼'].apply(lambda k: contract_mapping.get(k, {}).get('通路', ''))
            df_combined['課別'] = df_combined['合約號碼'].apply(lambda k: contract_mapping.get(k, {}).get('部門', ''))
            df_combined['報價單號'] = df_combined['合約號碼'].apply(lambda k: contract_mapping.get(k, {}).get('報價單號', ''))
            df_combined['匯率'] = df_combined['合約號碼'].apply(lambda k: contract_mapping.get(k, {}).get('匯率', ''))
            df_combined['業務員'] = df_combined['合約號碼'].apply(lambda k: contract_mapping.get(k, {}).get('業務', ''))
            df_combined['報價銅'] = df_combined['合約號碼'].apply(lambda k: contract_mapping.get(k, {}).get('報價銅價', ''))
            # 線種內容替換
            df_combined['線種'] = df_combined['線種'].replace({'銅通信電纜': '通信', '光通信電纜': '通信'})
            # 通路內容替換
            df_combined['通路'] = df_combined['通路'].replace({
                '電力': '經銷長約',
                '經銷專案-特定': '經銷專案',
                '電力專案': '專案',
                '經銷專案-產電': '經銷專案產電',
                '電力專案-產電': '經銷專案產電',
                '經銷專案-特開': '經銷專案',
                '經銷專案-新商模': '經銷專案',
                '經銷專案-綠能電力': '經銷專案',
                '經銷專案-類長約': '經銷專案'
            })
            # 計算報價銅成本 = 報價銅 * 銅量
            df_combined['報價銅'] = pd.to_numeric(df_combined['報價銅'], errors='coerce').fillna(0)
            df_combined['報價銅成本'] = df_combined['報價銅'] * df_combined['銅量']
        # 如果有產品群檔案，mapping 產品群
        if not df_product.empty:
            # 轉型料號為str確保匹配
            df_combined['物料'] = df_combined['物料'].astype(str)
            df_product['料號'] = df_product['料號'].astype(str)
            # 處理重複料號，取第一個
            df_product = df_product.drop_duplicates(subset='料號', keep='first')
            product_mapping = df_product.set_index('料號')['產品群'].to_dict()
            df_combined['產品群'] = df_combined['物料'].map(product_mapping).fillna('')
        # 如果有報價資訊檔案，覆寫報價銅和匯率
        if not df_quote.empty:
            df_quote['合約編號'] = df_quote['合約編號'].astype(str)
            df_combined['合約號碼'] = df_combined['合約號碼'].astype(str)
            df_quote = df_quote.drop_duplicates(subset='合約編號', keep='first')
            quote_mapping = df_quote.set_index('合約編號')[['銅價+銅價調整', '匯率']].to_dict(orient='index')
            # 有對到的覆寫，沒對到的保留原值
            quote_copper = df_combined['合約號碼'].apply(lambda k: quote_mapping.get(k, {}).get('銅價+銅價調整', None))
            quote_rate = df_combined['合約號碼'].apply(lambda k: quote_mapping.get(k, {}).get('匯率', None))
            df_combined['報價銅'] = quote_copper.where(quote_copper.notna(), df_combined.get('報價銅', ''))
            df_combined['匯率'] = quote_rate.where(quote_rate.notna(), df_combined.get('匯率', ''))
            # 覆寫後重新計算報價銅成本
            df_combined['報價銅'] = pd.to_numeric(df_combined['報價銅'], errors='coerce').fillna(0)
            df_combined['報價銅成本'] = df_combined['報價銅'] * df_combined['銅量']
        # 新增分類邏輯
        df_combined['分類'] = ''
        mask = df_combined['分類'] == ''
        # Step1 & Step2: 經銷長約(M-1) and (M-2)
        def classify_purchase_order(po, current_month):
            if isinstance(po, str) and '-1=Y' in po:
                parts = po.split(' ')
                if len(parts) > 1:
                    mm_part = parts[1].split('-1=')[0]
                    try:
                        mm = int(mm_part)
                        if mm == current_month:
                            return "經銷長約(M-1)"
                        elif mm < current_month:
                            return "經銷長約(M-2)"
                    except ValueError:
                        pass
            return ''
        df_combined.loc[mask, '分類'] = df_combined.loc[mask, '採購單'].apply(classify_purchase_order, args=(month,))
        # Step3: 銅量==0 填 "無"
        mask = df_combined['分類'] == ''
        df_combined.loc[mask & (df_combined['銅量'] == 0), '分類'] = '無'
        # Step4: 課別=="民電業務部/營業一課" 或 "民電業務部/營業三課" 填 "民電"
        mask = df_combined['分類'] == ''
        df_combined.loc[mask & ((df_combined['課別'] == '民電業務部/營業一課') | (df_combined['課別'] == '民電業務部/營業三課')), '分類'] = '民電'
        # Step5: 課別=="公電業務部/營業一課" 或 "公電業務部/營業二課" 填 "公電"
        mask = df_combined['分類'] == ''
        df_combined.loc[mask & ((df_combined['課別'] == '公電業務部/營業一課') | (df_combined['課別'] == '公電業務部/營業二課')), '分類'] = '公電'
        # Step6: 課別前兩個字=="產業" 或 "國際" 填 "外銷"
        mask = df_combined['分類'] == ''
        df_combined.loc[mask & ((df_combined['課別'].str[:2] == '產業') | (df_combined['課別'].str[:2] == '國際')), '分類'] = '外銷'
        # Step7: 線種=="通信" 填 "通信"
        mask = df_combined['分類'] == ''
        df_combined.loc[mask & (df_combined['線種'] == '通信'), '分類'] = '通信'

        # 新增訂單月邏輯
        def extract_mm(po):
            if isinstance(po, str) and '-1=Y' in po:
                po = po.replace(' ', '')  # 移除空格，提高格式彈性
                if '-1=Y' in po:
                    try:
                        mm_part = po.split('-1=Y')[0][-2:]  # 取最後兩個字符作為 mm
                        mm = int(mm_part)
                        if 1 <= mm <= 12:
                            return f"{mm:02d}"  # 返回字串
                    except ValueError:
                        pass
            return ''
        mask_m2 = df_combined['分類'] == "經銷長約(M-2)"
        df_combined.loc[mask_m2, '訂單月'] = df_combined.loc[mask_m2, '採購單'].apply(extract_mm)
        # 轉型訂單月為數字（匹配 m2_dict 的 int 鍵）
        df_combined['訂單月'] = pd.to_numeric(df_combined['訂單月'], errors='coerce')

        # 根據分類填報價銅
        df_combined.loc[df_combined['分類'] == "經銷長約(M-1)", '報價銅'] = m1_copper_price
        # 對於M-2，使用訂單月映射m2_dict
        mask_m2 = df_combined['分類'] == "經銷長約(M-2)"
        df_combined.loc[mask_m2, '報價銅'] = df_combined.loc[mask_m2, '訂單月'].map(m2_dict).fillna(df_combined.loc[mask_m2, '報價銅'])
        # 重新計算報價銅成本
        df_combined['報價銅成本'] = df_combined['報價銅'] * df_combined['銅量']

        # 重命名輸出欄位
        df_combined = df_combined.rename(columns={
            '參考文件號碼': '文件(Billing號)',
            '項目': 'billing項目',
            '銷售文件': '銷售文件',
            'sales_item': '銷售項目',
            '工廠': '工廠',
            '客戶': '客戶',
            '以 PCLC 計': '以 PCLC 計',
            '數量': '數量',
            '過帳日期': '過帳日期',
            'BUn': 'BUn'
        })
        # 過帳日期只抓年月日 (yyyy/mm/dd)
        df_combined['過帳日期'] = pd.to_datetime(df_combined['過帳日期'], errors='coerce').dt.strftime('%Y/%m/%d')
        # 確保所有輸出欄位存在，缺的設為空
        output_cols = ['文件(Billing號)', '物料', '品名', '產品群', '工廠', '線種', '課別', '通路', '客戶', '客戶名稱', '銷售文件', '銷售項目', 'billing項目', '以 PCLC 計', '數量', '過帳日期', 'BUn', '單位用銅', '銅量', '合約號碼', '採購單', '分類', '報價單號', '報價銅', '報價銅成本', '匯率', '訂單月', '業務員']
        for col in output_cols:
            if col not in df_combined.columns:
                df_combined[col] = ''
        # 選擇輸出欄位
        df_output = df_combined[output_cols]
        # 顯示預覽
        st.write("處理結果預覽（前10行）：")
        st.dataframe(df_output.head(10))
        # 匯出
        output = BytesIO()
        df_output.to_excel(output, index=False)
        output.seek(0)
        st.download_button(
            label="下載結果 Excel",
            data=output,
            file_name="mapped_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"錯誤：{str(e)}。請檢查檔案格式是否正確。")
