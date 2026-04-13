import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Sales Manager", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; background-color: #f8fafc; }
    .main { padding: 2rem 5rem; }
    h1 { font-weight: 800; color: #0f172a; letter-spacing: -0.05em; margin-bottom: 2rem; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px; }
    h3 { font-weight: 700; color: #1e293b; margin-top: 2rem; }
    [data-testid="stMetric"] { background-color: white; border: 1px solid #e2e8f0; padding: 25px !important; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
    [data-testid="stMetricValue"] { color: #0f172a; font-weight: 700 !important; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; background-color: #f1f5f9; padding: 8px; border-radius: 12px; }
    .stTabs [data-baseweb="tab"] { height: 45px; border-radius: 8px; border: none; padding: 0 30px; font-weight: 600; color: #64748b; }
    .stTabs [aria-selected="true"] { background-color: #0f172a !important; color: white !important; }
    .stDownloadButton > button { background-color: #0f172a !important; color: white !important; border-radius: 8px !important; width: 100%; padding: 15px; font-weight: 700; border: none !important; text-transform: uppercase; margin-top: 20px; }
    </style>
    """, unsafe_allow_html=True)

st.title("Sales Manager")

def clean_currency(value):
    if isinstance(value, str):
        return float(value.replace('$', '').replace(',', '').strip())
    return value

def calculate_diff(row):
    try:
        end = datetime(int(row['Năm Nhận File']), int(row['Tháng nhận file']), 1)
        if pd.isna(row['THÁNG NHẬN LEAD']) or pd.isna(row['NĂM NHẬN LEAD']): return None
        start = datetime(int(row['NĂM NHẬN LEAD']), int(row['THÁNG NHẬN LEAD']), 1)
        return (end.year - start.year) * 12 + (end.month - start.month)
    except: return None

uploaded_file = st.sidebar.file_uploader("Upload Data File", type=["xlsx"])

if uploaded_file:
    df_sales = pd.read_excel(uploaded_file, sheet_name=0)
    df_leads = pd.read_excel(uploaded_file, sheet_name=1)

    # 1. Logic Doanh số & Tốc độ
    df_sales['ANNUAL PREMIUM'] = df_sales['ANNUAL PREMIUM'].apply(clean_currency)
    df_sales['TARGET PREMIUM'] = df_sales['TARGET PREMIUM'].apply(clean_currency)
    df_sales['DOANH SỐ THỰC'] = df_sales[['ANNUAL PREMIUM', 'TARGET PREMIUM']].min(axis=1)
    df_sales['MONTHS_DIFF'] = df_sales.apply(calculate_diff, axis=1)
    df_sales['ĐÁNH GIÁ'] = df_sales['MONTHS_DIFF'].apply(lambda x: "Tự khai thác" if pd.isna(x) else (f"{int(x)} tháng - {'Chậm' if x > 6 else 'Nhanh'}"))

    # 2. Xử lý Tỉ lệ chuyển đổi Tổng (Cho file Xuất)
    total_leads_received = df_leads['LEAD ID'].nunique()
    total_leads_closed = df_sales[df_sales['SOURCE'] == 'SF']['LEAD ID'].nunique()
    total_conv_rate = (total_leads_closed / total_leads_received * 100) if total_leads_received > 0 else 0
    
    # 3. Xuất File Excel Đa Sheet & Màu sắc
    st.sidebar.markdown("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Sheet 1: Dashboard Tổng hợp
        summary_data = []
        for team in ['G', 'H', 'T']:
            team_closed = df_sales[(df_sales['TEAM'] == team) & (df_sales['SOURCE'] == 'SF')]['LEAD ID'].nunique()
            # Ước tính Lead nhận theo team (Dựa trên Owner trong Sheet 2)
            team_owners = df_sales[df_sales['TEAM'] == team]['OWNER'].unique()
            team_received = df_leads[df_leads['OWNER'].isin(team_owners)]['LEAD ID'].nunique()
            team_sales_val = df_sales[df_sales['TEAM'] == team]['DOANH SỐ THỰC'].sum()
            summary_data.append({
                'TEAM': team,
                'Tổng Lead Nhận': team_received,
                'Tổng Lead Chốt': team_closed,
                'Tỉ lệ Chuyển Đổi (%)': round((team_closed/team_received*100), 2) if team_received > 0 else 0,
                'Tổng Doanh Số ($)': team_sales_val
            })
        
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, sheet_name='DASHBOARD TỔNG', index=False)
        
        # Format Sheet Dashboard
        workbook = writer.book
        worksheet_dash = writer.sheets['DASHBOARD TỔNG']
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#0f172a', 'font_color': 'white', 'border': 1})
        num_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
        
        for col_num, value in enumerate(df_summary.columns.values):
            worksheet_dash.write(0, col_num, value, header_fmt)
            worksheet_dash.set_column(col_num, col_num, 20)

        # Sheet 2: Chi tiết dữ liệu
        df_sales.to_excel(writer, sheet_name='CHI TIẾT CHỐT', index=False)
        worksheet_detail = writer.sheets['CHI TIẾT CHỐT']
        red_fmt = workbook.add_format({'font_color': '#ef4444', 'bold': True})
        
        # Duyệt bảng để tô màu "Chậm"
        eval_col_idx = df_sales.columns.get_loc('ĐÁNH GIÁ')
        for row_num, val in enumerate(df_sales['ĐÁNH GIÁ']):
            if "Chậm" in str(val):
                worksheet_detail.write(row_num + 1, eval_col_idx, val, red_fmt)

    st.sidebar.download_button(
        label="DOWNLOAD FINAL REPORT",
        data=output.getvalue(),
        file_name=f"Manager_Report_{datetime.now().strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # --- HIỂN THỊ WEB APP ---
    tab1, tab2, tab3 = st.tabs(["TỔNG QUAN", "NGUỒN SF", "NGUỒN CC"])

    with tab1:
        st.metric("TỔNG DOANH SỐ TOÀN TEAM", f"${df_sales['DOANH SỐ THỰC'].sum():,.2f}")
        c1, c2 = st.columns([3, 2], gap="large")
        with c1:
            st.markdown("### Doanh số theo Team")
            st.bar_chart(df_summary.set_index('TEAM')['Tổng Doanh Số ($)'], color="#0f172a")
        with c2:
            st.markdown("### Tỉ lệ chuyển đổi 3 Team")
            st.dataframe(df_summary[['TEAM', 'Tỉ lệ Chuyển Đổi (%)']], use_container_width=True, hide_index=True)

    with tab2:
        df_sf = df_sales[df_sales['SOURCE'] == 'SF']
        st.markdown("### Hiệu quả chuyển đổi SF")
        df_leads['Month_Year'] = df_leads['DATE ADDED'].dt.strftime('%m/%Y')
        l_summary = df_leads.groupby(['OWNER', 'Month_Year'])['LEAD ID'].count().reset_index(name='Nhận')
        closed_ids = df_sf['LEAD ID'].unique()
        df_leads['Is_Closed'] = df_leads['LEAD ID'].isin(closed_ids)
        c_summary = df_leads[df_leads['Is_Closed'] == True].groupby(['OWNER', 'Month_Year'])['LEAD ID'].count().reset_index(name='Chốt')
        conv_final = pd.merge(l_summary, c_summary, on=['OWNER', 'Month_Year'], how='left').fillna(0)
        conv_final['%'] = (conv_final['Chốt'] / conv_final['Nhận'] * 100).round(2)
        st.dataframe(conv_final.sort_values(by=['Month_Year', '%'], ascending=False), use_container_width=True, hide_index=True)

        st.markdown("### Chi tiết xử lý SF")
        st.dataframe(
            df_sf[['OWNER', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC', 'ĐÁNH GIÁ']],
            use_container_width=True, hide_index=True,
            column_config={"DOANH SỐ THỰC": st.column_config.NumberColumn(format="$%.2f")}
        )

    with tab3:
        df_cc = df_sales[df_sales['SOURCE'] == 'CC']
        st.markdown("### Chi tiết chốt CC (Tự khai thác)")
        st.dataframe(
            df_cc[['OWNER', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC']], 
            use_container_width=True, hide_index=True,
            column_config={"DOANH SỐ THỰC": st.column_config.NumberColumn(format="$%.2f")}
        )
else:
    st.info("Anh hãy tải file lên để bắt đầu soi số liệu.")
