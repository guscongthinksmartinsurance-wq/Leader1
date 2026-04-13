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
    h3 { font-weight: 700; color: #1e293b; margin-top: 1.5rem; }
    [data-testid="stMetric"] { background-color: white; border: 1px solid #e2e8f0; padding: 25px !important; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
    [data-testid="stMetricValue"] { color: #0f172a; font-weight: 700 !important; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; background-color: #f1f5f9; padding: 8px; border-radius: 12px; }
    .stTabs [data-baseweb="tab"] { height: 45px; border-radius: 8px; border: none; padding: 0 30px; font-weight: 600; color: #64748b; }
    .stTabs [aria-selected="true"] { background-color: #0f172a !important; color: white !important; }
    .stDownloadButton > button { background-color: #0f172a !important; color: white !important; border-radius: 8px !important; width: 100%; padding: 15px; font-weight: 700; border: none !important; text-transform: uppercase; }
    </style>
    """, unsafe_allow_html=True)

st.title("Sales Manager")

def clean_currency(value):
    if isinstance(value, str):
        return float(value.replace('$', '').replace(',', '').strip())
    return value

def calculate_diff(row):
    try:
        end = datetime(int(row['NĂM NHẬN FILE']), int(row['THÁNG NHẬN FILE']), 1)
        if pd.isna(row['THÁNG NHẬN LEAD']) or pd.isna(row['NĂM NHẬN LEAD']): return None
        start = datetime(int(row['NĂM NHẬN LEAD']), int(row['THÁNG NHẬN LEAD']), 1)
        return (end.year - start.year) * 12 + (end.month - start.month)
    except: return None

uploaded_file = st.sidebar.file_uploader("Upload Data File (2 Sheets)", type=["xlsx"])

if uploaded_file:
    df_sales = pd.read_excel(uploaded_file, sheet_name=0)
    df_leads = pd.read_excel(uploaded_file, sheet_name=1)

    df_sales.columns = df_sales.columns.str.strip().str.upper()
    df_leads.columns = df_leads.columns.str.strip().str.upper()

    df_sales['LEAD ID'] = df_sales['LEAD ID'].astype(str)
    df_leads['LEAD ID'] = df_leads['LEAD ID'].astype(str)

    df_sales['ANNUAL PREMIUM'] = df_sales['ANNUAL PREMIUM'].apply(clean_currency)
    df_sales['TARGET PREMIUM'] = df_sales['TARGET PREMIUM'].apply(clean_currency)
    df_sales['DOANH SỐ THỰC'] = df_sales[['ANNUAL PREMIUM', 'TARGET PREMIUM']].min(axis=1)
    df_sales['MONTHS_DIFF'] = df_sales.apply(calculate_diff, axis=1)
    
    def get_bucket(m):
        if pd.isna(m): return "Tự khai thác"
        if m <= 3: return "0-3 Tháng"
        if m <= 6: return "3-6 Tháng"
        if m <= 9: return "6-9 Tháng"
        if m <= 12: return "9-12 Tháng"
        return "> 12 Tháng"
    df_sales['THỜI GIAN CHỐT'] = df_sales['MONTHS_DIFF'].apply(get_bucket)

    # --- LOGIC ĐỐI SOÁT (SF ONLY) ---
    df_sf_only = df_sales[df_sales['SOURCE'] == 'SF']
    closed_ids = df_sf_only['LEAD ID'].unique()
    df_leads['IS_CLOSED'] = df_leads['LEAD ID'].isin(closed_ids)

    # 1. Chuyển đổi Team (Dựa trên SF)
    team_recv = df_leads.groupby('TEAM')['LEAD ID'].nunique().reset_index(name='Nhận')
    team_cls = df_leads[df_leads['IS_CLOSED']].groupby('TEAM')['LEAD ID'].nunique().reset_index(name='Chốt')
    # Doanh số SF theo Team
    team_sales_sf = df_sf_only.groupby('TEAM')['DOANH SỐ THỰC'].sum().reset_index(name='Doanh số SF')
    df_team_final = pd.merge(team_recv, team_cls, on='TEAM', how='left').fillna(0)
    df_team_final = pd.merge(df_team_final, team_sales_sf, on='TEAM', how='left').fillna(0)
    df_team_final['% Chuyển đổi'] = (df_team_final['Chốt'] / df_team_final['Nhận'] * 100).round(2)

    # 2. Hiệu suất Nhân viên (Tháng & Tích lũy)
    date_col = [c for c in df_leads.columns if 'DATE' in c][0]
    df_leads['THÁNG'] = df_leads[date_col].dt.month
    l_month = df_leads.groupby(['OWNER', 'TEAM', 'THÁNG']).size().reset_index(name='Nhận_T')
    c_month = df_leads[df_leads['IS_CLOSED']].groupby(['OWNER', 'THÁNG']).size().reset_index(name='Chốt_T')
    perf_month = pd.merge(l_month, c_month, on=['OWNER', 'THÁNG'], how='left').fillna(0)
    
    l_tot = df_leads.groupby('OWNER').size().reset_index(name='Tổng_Nhận')
    c_tot = df_leads[df_leads['IS_CLOSED']].groupby('OWNER').size().reset_index(name='Tổng_Chốt')
    perf_tot = pd.merge(l_tot, c_tot, on='OWNER', how='left').fillna(0)
    perf_tot['Tổng Tỉ Lệ (%)'] = (perf_tot['Tổng_Chốt'] / perf_tot['Tổng_Nhận'] * 100).round(2)
    perf_final = pd.merge(perf_month, perf_tot[['OWNER', 'Tổng Tỉ Lệ (%)']], on='OWNER', how='left')

    # 3. Doanh số CC & Sản phẩm
    df_cc_only = df_sales[df_sales['SOURCE'] == 'CC']
    total_sales_cc = df_cc_only['DOANH SỐ THỰC'].sum()
    total_sales_sf = df_sf_only['DOANH SỐ THỰC'].sum()
    prod_summary = df_sales.groupby('PRODUCT').size().reset_index(name='Qty')

    # --- XUẤT FILE EXCEL QUẢN TRỊ ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        h_fmt = workbook.add_format({'bold': True, 'bg_color': '#0f172a', 'font_color': 'white', 'border': 1, 'align': 'center'})
        red_fmt = workbook.add_format({'font_color': '#ef4444', 'bold': True, 'border': 1})
        num_fmt = workbook.add_format({'num_format': '$#,##0', 'border': 1})

        # Sheet DASHBOARD
        df_team_final.to_excel(writer, sheet_name='DASHBOARD', index=False, startrow=2)
        ws_dash = writer.sheets['DASHBOARD']
        ws_dash.write('A1', f'BÁO CÁO FUNNEL - CHUYỂN ĐỔI CHUNG: {round(df_team_final["Chốt"].sum()/df_team_final["Nhận"].sum()*100, 2)}%', workbook.add_format({'bold': True, 'font_size': 14}))
        for i, col in enumerate(df_team_final.columns): ws_dash.write(2, i, col, h_fmt)
        
        c_sales = workbook.add_chart({'type': 'column'})
        c_sales.add_series({'categories': '=DASHBOARD!$A$4:$A$6', 'values': '=DASHBOARD!$D$4:$D$6', 'name': 'Doanh số SF', 'fill': {'color': '#1e293b'}})
        ws_dash.insert_chart('F2', c_sales)

        # Sheet NHÂN VIÊN
        perf_final.to_excel(writer, sheet_name='NHAN_VIEN', index=False)
        ws_emp = writer.sheets['NHAN_VIEN']
        ws_emp.set_column('A:G', 18)
        for i, col in enumerate(perf_final.columns): ws_emp.write(0, i, col, h_fmt)

        # Sheet CHI TIẾT
        df_sales.to_excel(writer, sheet_name='CHI_TIET_DATA', index=False)
        ws_det = writer.sheets['CHI_TIET_DATA']
        ws_det.set_column('A:Z', 15)
        eval_idx = df_sales.columns.get_loc('THỜI GIAN CHỐT')
        for r, v in enumerate(df_sales['THỜI GIAN CHỐT']):
            if "Chậm" in str(v) or "> 12" in str(v): ws_det.write(r+1, eval_idx, v, red_fmt)

        # Sheet SẢN PHẨM
        prod_summary.to_excel(writer, sheet_name='SAN_PHAM', index=False)
        ws_p = writer.sheets['SAN_PHAM']
        c_pie = workbook.add_chart({'type': 'pie'})
        c_pie.add_series({'categories': '=SAN_PHAM!$A$2:$A$10', 'values': '=SAN_PHAM!$B$2:$B$10', 'name': 'Tỉ lệ Sản phẩm'})
        ws_p.insert_chart('D2', c_pie)

    st.sidebar.download_button("Download Final Report", data=output.getvalue(), file_name=f"Manager_Report_{datetime.now().strftime('%d%m%Y')}.xlsx")

    # --- UI WEB ---
    t1, t2, t3 = st.tabs(["📊 TỔNG QUAN", "👥 HIỆU SUẤT SALE", "🔍 DỮ LIỆU CHI TIẾT"])
    
    with t1:
        c1, c2, c3 = st.columns(3)
        c1.metric("TỔNG DOANH SỐ (SF+CC)", f"${(total_sales_sf + total_sales_cc):,.2f}")
        c2.metric("TỔNG NHẬN SF", f"{df_leads['LEAD ID'].nunique():,}")
        c3.metric("CHUYỂN ĐỔI SF CHUNG", f"{round(df_team_final['Chốt'].sum()/df_team_final['Nhận'].sum()*100, 2)}%")
        
        st.markdown("### Hiệu suất 3 Team (Chỉ nguồn SF)")
        st.dataframe(df_team_final, use_container_width=True, hide_index=True)
        
        st.markdown("### Thống kê Sản phẩm Tổng (SF + CC)")
        st.dataframe(prod_summary.sort_values(by='Số lượng chốt', ascending=False), use_container_width=True, hide_index=True)

    with t2:
        st.markdown("### Chuyển đổi Nhân viên (Tháng & Tổng tích lũy)")
        st.dataframe(perf_final.sort_values(by=['THÁNG', 'Tổng Tỉ Lệ (%)'], ascending=[True, False]), use_container_width=True, hide_index=True)

    with t3:
        st.markdown("### Toàn bộ lịch sử chốt khách")
        st.dataframe(df_sales[['OWNER', 'TEAM', 'LEAD ID', 'PRODUCT', 'SOURCE', 'DOANH SỐ THỰC', 'THỜI GIAN CHỐT']], use_container_width=True, hide_index=True)
else:
    st.info("Anh Công hãy tải file Excel có cột TEAM ở Sheet 2 để bắt đầu.")
