import streamlit as st
import pandas as pd
from io import BytesIO
import datetime

st.title("Excel 銷售日報表 Mapping 工具")

tab_daily, tab_monthly = st.tabs(["📋 每日作業", "📋 月底作業"])

# ===== 每日作業 =====
with tab_daily:
    st.subheader("每日 Mapping")
    st.caption("上傳原始檔案進行 mapping，匯出後可自行手動調整並堆疊累積。")

    sales_file = st.file_uploader("41110000 銷貨收入", type="xlsx", key="d_sales")
    returns_file = st.file_uploader("41700000 銷貨退回 (選填)", type="xlsx", key="d_returns")
    zsdc_file = st.file_uploader("zsdc", type="xlsx", key="d_zsdc")
    contract_file = st.file_uploader("合約管理", type="xlsx", key="d_contract")
    product_file = st.file_uploader("產品群", type="xlsx", key="d_product")

    current_month = datetime.datetime.now().month
    month = st.number_input("設定月份 (預設當月)", min_value=1, max_value=12, value=current_month, key="d_month")
    start_date = st.date_input("開始日期 (篩選輸入日期)", value=datetime.datetime.now().date() - datetime.timedelta(days=7), key="d_start")
    end_date = st.date_input("結束日期 (篩選輸入日期)", value=datetime.datetime.now().date(), key="d_end")

    if st.button("處理檔案", key="btn_daily") and sales_file and zsdc_file:
        try:
            df_sales = pd.read_excel(sales_file, sheet_name='工作表1')
            df_zsdc = pd.read_excel(zsdc_file, sheet_name='工作表1')
            if returns_file:
                df_returns = pd.read_excel(returns_file, sheet_name='工作表1')
            else:
                df_returns = pd.DataFrame()
            if contract_file:
                df_contract = pd.read_excel(contract_file, sheet_name='工作表1')
            else:
                df_contract = pd.DataFrame()
            if product_file:
                df_product = pd.read_excel(product_file, sheet_name='工作表1')
            else:
                df_product = pd.DataFrame()

            # 篩選輸入日期
            df_sales['輸入日期'] = pd.to_datetime(df_sales['輸入日期'], errors='coerce')
            df_sales = df_sales[(df_sales['輸入日期'] >= pd.to_datetime(start_date)) & (df_sales['輸入日期'] <= pd.to_datetime(end_date))]
            if not df_returns.empty:
                df_returns['輸入日期'] = pd.to_datetime(df_returns['輸入日期'], errors='coerce')
                df_returns = df_returns[(df_returns['輸入日期'] >= pd.to_datetime(start_date)) & (df_returns['輸入日期'] <= pd.to_datetime(end_date))]

            df_sales = df_sales.rename(columns={'項目.1': 'sales_item'})
            if not df_returns.empty:
                df_returns = df_returns.rename(columns={'項目.1': 'sales_item'})

            common_cols = ['參考文件號碼', '項目', '物料', '工廠', '客戶', '銷售文件', 'sales_item', '以 PCLC 計', '數量', '過帳日期', 'BUn', '輸入日期']
            df_sales_common = df_sales[common_cols].copy()
            df_sales_common['以 PCLC 計'] = -df_sales_common['以 PCLC 計']
            df_sales_common['數量'] = -df_sales_common['數量']
            if not df_returns.empty:
                df_returns_common = df_returns[common_cols].copy()
                df_returns_common['以 PCLC 計'] = -df_returns_common['以 PCLC 計']
                df_returns_common['數量'] = -df_returns_common['數量']
            else:
                df_returns_common = pd.DataFrame(columns=common_cols)

            df_combined = pd.concat([df_sales_common, df_returns_common], ignore_index=True)

            df_combined['銷售文件'] = pd.to_numeric(df_combined['銷售文件'], errors='coerce').fillna(0).astype(int)
            df_zsdc['先前文件'] = pd.to_numeric(df_zsdc['先前文件'], errors='coerce').fillna(0).astype(int)
            df_combined['參考文件號碼'] = pd.to_numeric(df_combined['參考文件號碼'], errors='coerce').fillna(0).astype(int)
            df_combined['項目'] = pd.to_numeric(df_combined['項目'], errors='coerce').fillna(0).astype(int)
            df_combined['key'] = df_combined['參考文件號碼'].astype(str) + '_' + df_combined['項目'].astype(str)
            df_zsdc['文件'] = pd.to_numeric(df_zsdc['文件'], errors='coerce').fillna(0).astype(int)
            df_zsdc['項目'] = pd.to_numeric(df_zsdc['項目'], errors='coerce').fillna(0).astype(int)
            df_zsdc['key'] = df_zsdc['文件'].astype(str) + '_' + df_zsdc['項目'].astype(str)
            df_zsdc = df_zsdc.drop_duplicates(subset='key', keep='first')

            mapping_dict = df_zsdc.set_index('key')[['物料說明', '淨重']].to_dict(orient='index')
            df_combined['品名'] = df_combined['key'].apply(lambda k: mapping_dict.get(k, {}).get('物料說明', ''))
            df_combined['單位用銅'] = df_combined['key'].apply(lambda k: mapping_dict.get(k, {}).get('淨重', ''))

            # 記住沒有 zsdc 對應的行（主要是銷貨退回），後續輸出前清空數值欄位
            no_zsdc_mask = df_combined['單位用銅'] == ''

            sales_doc_dict = df_zsdc.set_index('先前文件')['bill-to-name'].to_dict()
            df_combined['客戶名稱'] = df_combined['銷售文件'].map(sales_doc_dict).fillna('')
            contract_no_dict = df_zsdc.set_index('先前文件')['合約號碼'].to_dict()
            purchase_order_dict = df_zsdc.set_index('先前文件')['採購單號碼'].to_dict()
            df_combined['合約號碼'] = df_combined['銷售文件'].map(contract_no_dict).fillna('')
            df_combined['採購單'] = df_combined['銷售文件'].map(purchase_order_dict).fillna('')

            df_combined['單位用銅'] = pd.to_numeric(df_combined['單位用銅'], errors='coerce').fillna(0)
            df_combined['銅量'] = df_combined['單位用銅'] * df_combined['數量']

            if not df_contract.empty:
                df_combined['合約號碼'] = df_combined['合約號碼'].astype(str)
                df_contract['合約編號'] = df_contract['合約編號'].astype(str)
                df_contract = df_contract.drop_duplicates(subset='合約編號', keep='first')
                contract_mapping = df_contract.set_index('合約編號')[['產品部', '通路', '部門', '報價單號', '匯率', '業務', '報價銅價']].to_dict(orient='index')
                df_combined['線種'] = df_combined['合約號碼'].apply(lambda k: contract_mapping.get(k, {}).get('產品部', ''))
                df_combined['產品群'] = ''
                df_combined['通路'] = df_combined['合約號碼'].apply(lambda k: contract_mapping.get(k, {}).get('通路', ''))
                df_combined['課別'] = df_combined['合約號碼'].apply(lambda k: contract_mapping.get(k, {}).get('部門', ''))
                df_combined['報價單號'] = df_combined['合約號碼'].apply(lambda k: contract_mapping.get(k, {}).get('報價單號', ''))
                df_combined['匯率'] = df_combined['合約號碼'].apply(lambda k: contract_mapping.get(k, {}).get('匯率', ''))
                df_combined['業務員'] = df_combined['合約號碼'].apply(lambda k: contract_mapping.get(k, {}).get('業務', ''))
                df_combined['報價銅'] = df_combined['合約號碼'].apply(lambda k: contract_mapping.get(k, {}).get('報價銅價', ''))
                df_combined['線種'] = df_combined['線種'].replace({'銅通信電纜': '通信', '光通信電纜': '通信'})
                df_combined['通路'] = df_combined['通路'].replace({
                    '電力': '經銷長約',
                    '經銷專案-特定': '經銷專案',
                    '電力專案': '專案',
                    '經銷專案-產電': '經銷專案產電',
                    '電力專案-產電': '經銷專案產電',
                    '經銷專案-特開': '經銷專案',
                    '經銷專案-新商模': '經銷專案',
                    '經銷專案-綠能電力': '經銷專案',
                    '經銷專案-類長約': '經銷專案',
                    '電力專案-類長約': '經銷專案',
                    '通信專案': '民間通信',
                    '通信': '民間通信',
                })
                df_combined['報價銅'] = pd.to_numeric(df_combined['報價銅'], errors='coerce').fillna(0)
                df_combined['報價銅成本'] = df_combined['報價銅'] * df_combined['銅量']

            if not df_product.empty:
                df_combined['物料'] = df_combined['物料'].astype(str)
                df_product['料號'] = df_product['料號'].astype(str)
                df_product = df_product.drop_duplicates(subset='料號', keep='first')
                product_mapping = df_product.set_index('料號')['產品群'].to_dict()
                df_combined['產品群'] = df_combined['物料'].map(product_mapping).fillna('')
                # 品名為空時（如銷貨退回），從產品群檔案補品名
                if '品名' in df_product.columns:
                    product_name_mapping = df_product.set_index('料號')['品名'].to_dict()
                    empty_mask = df_combined['品名'] == ''
                    df_combined.loc[empty_mask, '品名'] = df_combined.loc[empty_mask, '物料'].map(product_name_mapping).fillna('')

            # 分類邏輯
            df_combined['分類'] = ''
            mask = df_combined['分類'] == ''

            def classify_purchase_order(po, current_month):
                if isinstance(po, str) and '-1=Y' in po:
                    parts = po.split(' ')
                    if len(parts) > 1:
                        mm_part = parts[1].split('-1=')[0]
                        try:
                            mm = int(mm_part)
                            if mm == current_month:
                                return "經銷長約(M-1)"
                            else:
                                return "經銷長約(M-2)"
                        except ValueError:
                            pass
                return ''

            df_combined.loc[mask, '分類'] = df_combined.loc[mask, '採購單'].apply(classify_purchase_order, args=(month,))
            mask = df_combined['分類'] == ''
            df_combined.loc[mask & (df_combined['銅量'] == 0), '分類'] = '無'
            mask = df_combined['分類'] == ''
            民電_mask = mask & ((df_combined['課別'] == '民電業務部/營業一課') | (df_combined['課別'] == '民電業務部/營業三課'))
            df_combined.loc[民電_mask & (df_combined['線種'] == '通信'), '分類'] = '通信'
            df_combined.loc[民電_mask & (df_combined['線種'] != '通信'), '分類'] = '民電'
            mask = df_combined['分類'] == ''
            公電_mask = mask & ((df_combined['課別'] == '公電業務部/營業一課') | (df_combined['課別'] == '公電業務部/營業二課'))
            df_combined.loc[公電_mask & (df_combined['工廠'].str.strip() == 'SCDM'), '分類'] = '無'
            df_combined.loc[公電_mask & (df_combined['工廠'].str.strip() != 'SCDM') & (df_combined['線種'] == '通信'), '分類'] = '通信'
            df_combined.loc[公電_mask & (df_combined['工廠'].str.strip() != 'SCDM') & (df_combined['線種'] != '通信'), '分類'] = '公電'
            mask = df_combined['分類'] == ''
            df_combined.loc[mask & ((df_combined['課別'].str[:2] == '產業') | (df_combined['課別'].str[:2] == '國際')), '分類'] = '外銷'
            mask = df_combined['分類'] == ''
            df_combined.loc[mask & (df_combined['線種'] == '通信'), '分類'] = '通信'

            def extract_mm(po):
                if isinstance(po, str) and '-1=Y' in po:
                    po = po.replace(' ', '')
                    if '-1=Y' in po:
                        try:
                            mm_part = po.split('-1=Y')[0][-2:]
                            mm = int(mm_part)
                            if 1 <= mm <= 12:
                                return f"{mm:02d}"
                        except ValueError:
                            pass
                return ''

            mask_m2 = df_combined['分類'] == "經銷長約(M-2)"
            df_combined.loc[mask_m2, '訂單月'] = df_combined.loc[mask_m2, '採購單'].apply(extract_mm)
            df_combined['訂單月'] = pd.to_numeric(df_combined['訂單月'], errors='coerce')

            df_combined['報價銅'] = pd.to_numeric(df_combined.get('報價銅', 0), errors='coerce').fillna(0)
            df_combined['報價銅成本'] = df_combined['報價銅'] * df_combined['銅量']

            # 沒有 zsdc 對應的行（如銷貨退回），數值欄位清為空值，產品群和品名保留
            df_combined.loc[no_zsdc_mask, '單位用銅'] = ''
            df_combined.loc[no_zsdc_mask, '銅量'] = ''
            df_combined.loc[no_zsdc_mask, '報價銅'] = ''
            df_combined.loc[no_zsdc_mask, '報價銅成本'] = ''
            df_combined.loc[no_zsdc_mask, '匯率'] = ''
            df_combined.loc[no_zsdc_mask, '分類'] = ''

            df_combined = df_combined.rename(columns={
                '參考文件號碼': '文件(Billing號)',
                '項目': 'billing項目',
                'sales_item': '銷售項目',
            })
            df_combined['過帳日期'] = pd.to_datetime(df_combined['過帳日期'], errors='coerce').dt.strftime('%Y/%m/%d')
            output_cols = ['文件(Billing號)', '物料', '品名', '產品群', '工廠', '線種', '課別', '通路', '客戶', '客戶名稱', '銷售文件', '銷售項目', 'billing項目', '以 PCLC 計', '數量', '過帳日期', 'BUn', '單位用銅', '銅量', '合約號碼', '採購單', '分類', '訂單月', '報價銅', '報價銅成本', '匯率', '報價單號', '業務員']
            for col in output_cols:
                if col not in df_combined.columns:
                    df_combined[col] = ''
            df_output = df_combined[output_cols]

            st.write("處理結果預覽（前10行）：")
            st.dataframe(df_output.head(10))
            output = BytesIO()
            df_output.to_excel(output, index=False)
            output.seek(0)
            st.download_button(
                label="下載結果 Excel",
                data=output,
                file_name="mapped_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_daily"
            )
        except Exception as e:
            st.error(f"錯誤：{str(e)}。請檢查檔案格式是否正確。")

# ===== 月底作業 =====
with tab_monthly:
    st.subheader("月底作業")
    st.caption("上傳堆疊好的累積檔案 + 報價資訊，只覆寫報價銅、匯率及 M-1/M-2 銅價，不會動到你手動調整過的內容。")

    accumulated_file = st.file_uploader("上傳堆疊好的累積檔案", type="xlsx", key="m_accumulated")
    quote_file = st.file_uploader("報價資訊", type="xlsx", key="m_quote")

    current_month_m = datetime.datetime.now().month
    month_m = st.number_input("設定月份 (預設當月)", min_value=1, max_value=12, value=current_month_m, key="m_month")
    m1_copper_price = st.number_input("M-1銅價", value=0.00, step=0.01, format="%0.2f", key="m_m1")

    st.write("M-2銅價組 (可新增多組)")
    if 'm2_groups' not in st.session_state:
        st.session_state.m2_groups = [{'month': current_month_m - 2 if current_month_m > 2 else 10, 'price': 0.00}]
    if st.button("新增 M-2 組", key="btn_add_m2"):
        st.session_state.m2_groups.append({'month': current_month_m - 2 if current_month_m > 2 else 10, 'price': 0.00})
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

    if st.button("處理檔案（月底）", key="btn_monthly") and accumulated_file:
        try:
            # 直接讀取累積檔案，不重新 mapping
            df = pd.read_excel(accumulated_file)

            # 報價資訊覆寫報價銅和匯率（有對到才覆寫）
            if quote_file:
                df_quote = pd.read_excel(quote_file, sheet_name='qry_Temp')
                df_quote['合約編號'] = df_quote['合約編號'].astype(str)
                df['合約號碼'] = df['合約號碼'].astype(str)
                df_quote = df_quote.drop_duplicates(subset='合約編號', keep='first')
                quote_mapping = df_quote.set_index('合約編號')[['銅價+銅價調整', '匯率']].to_dict(orient='index')

                quote_copper = df['合約號碼'].apply(lambda k: quote_mapping.get(k, {}).get('銅價+銅價調整', None))
                quote_rate = df['合約號碼'].apply(lambda k: quote_mapping.get(k, {}).get('匯率', None))
                df['報價銅'] = quote_copper.where(quote_copper.notna(), df['報價銅'])
                df['匯率'] = quote_rate.where(quote_rate.notna(), df['匯率'])

            # M-1/M-2 銅價覆寫
            df['報價銅'] = pd.to_numeric(df['報價銅'], errors='coerce').fillna(0)
            df['銅量'] = pd.to_numeric(df['銅量'], errors='coerce').fillna(0)
            df['訂單月'] = pd.to_numeric(df['訂單月'], errors='coerce')

            df.loc[df['分類'] == "經銷長約(M-1)", '報價銅'] = m1_copper_price
            mask_m2 = df['分類'] == "經銷長約(M-2)"
            df.loc[mask_m2, '報價銅'] = df.loc[mask_m2, '訂單月'].map(m2_dict).fillna(df.loc[mask_m2, '報價銅'])

            # 重新計算報價銅成本
            df['報價銅'] = pd.to_numeric(df['報價銅'], errors='coerce').fillna(0)
            df['報價銅成本'] = df['報價銅'] * df['銅量']

            st.write("處理結果預覽（前10行）：")
            st.dataframe(df.head(10))
            output = BytesIO()
            df.to_excel(output, index=False)
            output.seek(0)
            st.download_button(
                label="下載結果 Excel（月底）",
                data=output,
                file_name="mapped_report_final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_monthly"
            )
        except Exception as e:
            st.error(f"錯誤：{str(e)}。請檢查檔案格式是否正確。")
