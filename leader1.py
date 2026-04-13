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
    .stTabs [data-baseweb="tab-list"] { gap: 10px; background-color: #f1f5f9; padding: 8px; border-radius: 12px; }
    .stTabs [data-baseweb="tab"] { height: 45px; border-radius: 8px; border: none; padding: 0 30px; font-weight: 600; color: #64748b; }
    .stTabs [aria-selected="true"] { background-color: #0f172a !important; color: white !important; }
    .stDownloadButton > button { background-color: #0f172a !important; color: white !important; border-radius: 8px !important; width: 100%; padding: 15px; font-weight: 700; border: none !important; text-transform: uppercase; }
    .stDataFrame { border: 1px solid #e2e8f0; border-radius: 10px; }
    </style>
    """, unsafe_allow_html=True)

st.title("Sales Manager")

def clean_currency(value):
    if isinstance(value, str):
        return float(value.replace('$', '').replace(',', '').strip())
    return value

uploaded_file = st.sidebar.file_uploader("Upload Data File (SF Only Logic)", type=["xlsx"])

if uploaded_file:
    # 1. Đọc và Chuẩn hóa
    df_sales_raw = pd.read_excel(uploaded_file, sheet_name=0)
    df_leads_raw = pd.read_excel(uploaded_file, sheet_name=1)
    
    df_sales_raw.columns = df_sales_raw.columns.str.strip().str.upper()
    df_leads_raw.columns = df_leads_raw.columns.str.strip().str.upper()
    
    df_sales_raw['LEAD ID'] = df_sales_raw['LEAD ID'].astype(str)
    df_leads_raw['LEAD ID'] = df_leads_raw['LEAD ID'].astype(str)
    
    # CHỈ LẤY NGUỒN SF
    df_sales = df_sales_raw[df_sales_raw['SOURCE'] == 'SF'].copy()
    df_leads = df_leads_raw.copy()

    df_sales['ANNUAL PREMIUM'] = df_sales['ANNUAL PREMIUM'].apply(clean_currency)
    df_sales['TARGET PREMIUM'] = df_sales['TARGET PREMIUM'].apply(clean_currency)
    df_sales['DOANH SỐ THỰC'] = df_sales[['ANNUAL PREMIUM', 'TARGET PREMIUM']].min(axis=1)

    # 2. Logic Chốt Nóng & Tỉ Lệ Tích Lũy
    df_leads['THÁNG_NHẬN'] = df_leads['DATE ADDED'].dt.month
    df_hot = pd.merge(
        df_leads[['OWNER', 'TEAM', 'LEAD ID', 'THÁNG_NHẬN']], 
        df_sales[['LEAD ID', 'THÁNG NHẬN FILE', 'PRODUCT', 'DOANH SỐ THỰC']], 
        on='LEAD ID', how='inner'
    )
    df_hot_closed = df_hot[df_hot['THÁNG_NHẬN'] == df_hot['THÁNG NHẬN FILE']]

    # Bảng chi tiết từng nhân viên (Theo tháng + Tổng tích lũy)
    e_recv = df_leads.groupby(['OWNER', 'THÁNG_NHẬN']).size().reset_index(name='Nhận_Tháng')
    e_hot = df_hot_closed.groupby(['OWNER', 'THÁNG_NHẬN']).size().reset_index(name='Chốt_Nóng')
    df_perf = pd.merge(e_recv, e_hot, on=['OWNER', 'THÁNG_NHẬN'], how='left').fillna(0)
    
    # Tính Tổng tích lũy 3 tháng của từng người
    l_tot = df_leads.groupby('OWNER').size().reset_index(name='Tổng_Nhận')
    c_tot = df_sales.groupby('OWNER').size().reset_index(name='Tổng_Chốt')
    perf_tot = pd.merge(l_tot, c_tot, on='OWNER', how='left').fillna(0)
    perf_tot['Tổng Tỉ Lệ (%)'] = (perf_tot['Tổng_Chốt'] / perf_tot['Tổng_Nhận'] * 100).round(2)
    
    df_emp_final = pd.merge(df_perf, perf_tot[['OWNER', 'Tổng Tỉ Lệ (%)']], on='OWNER', how='left')

    # Bảng Sản phẩm SF (Số lượng & Doanh số)
    df_prod_sf = df_sales.groupby('PRODUCT').agg({'LEAD ID': 'count', 'DOANH SỐ THỰC': 'sum'}).reset_index()
    df_prod_sf.columns = ['Sản phẩm', 'Số lượng chốt', 'Doanh số ($)']
    df_prod_sf = df_prod_sf.sort_values(by='Số lượng chốt', ascending=False)

    # --- XUẤT FILE EXCEL QUẢN TRỊ (SF ONLY) ---
    st.sidebar.markdown("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        # Formats
        h_fmt = workbook.add_format({'bold': True, 'bg_color': '#0f172a', 'font_color': 'white', 'border': 1, 'align': 'center'})
        money_fmt = workbook.add_format({'num_format': '$#,##0', 'border': 1})
        pct_fmt = workbook.add_format({'num_format': '0.00"%"', 'border': 1, 'bold': True})

        # Sheet 1: NHÂN VIÊN (Tháng & Tổng)
        df_emp_final.to_excel(writer, sheet_name='NHAN_VIEN_SF', index=False)
        ws_emp = writer.sheets['NHAN_VIEN_SF']
        ws_emp.set_column('A:G', 18)
        for i, col in enumerate(df_emp_final.columns): ws_emp.write(0, i, col, h_fmt)

        # Sheet 2: SẢN PHẨM SF (Kèm Biểu đồ & Doanh số)
        df_prod_sf.to_excel(writer, sheet_name='SAN_PHAM_SF', index=False)
        ws_p = writer.sheets['SAN_PHAM_SF']
        ws_p.set_column('A:C', 20)
        for i, col in enumerate(df_prod_sf.columns): ws_p.write(0, i, col, h_fmt)
        
        # Chèn biểu đồ tròn sản phẩm
        chart_p = workbook.add_chart({'type': 'pie'})
        chart_p.add_series({
            'name': 'Cơ cấu Sản phẩm',
            'categories': f'=SAN_PHAM_SF!$A$2:$A${len(df_prod_sf)+1}',
            'values': f'=SAN_PHAM_SF!$B$2:$B${len(df_prod_sf)+1}',
        })
        chart_p.set_title({'name': 'Tỉ lệ Sản phẩm theo Số lượng'})
        ws_p.insert_chart('E2', chart_p)

    st.sidebar.download_button("Download Final SF Report", data=output.getvalue(), file_name="Henry_SF_Manager_Report.xlsx")

    # --- UI WEB APP ---
    tab1, tab2 = st.tabs(["📊 BÁO CÁO TỔNG HỢP", "👥 CHI TIẾT NHÂN VIÊN"])
    
    with tab1:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("### Sản phẩm SF chủ lực")
            st.dataframe(df_prod_sf, use_container_width=True, hide_index=True)
        with c2:
            st.markdown("### Cơ cấu doanh số SF")
            st.bar_chart(df_prod_sf.set_index('Sản phẩm')['Doanh số ($)'], color="#0f172a")

        st.markdown("### Tổng hợp hiệu suất Funnel theo tháng")
        m_summary = df_emp_final.groupby('THÁNG_NHẬN').agg({'Nhận_Tháng': 'sum', 'Chốt_Nóng': 'sum'}).reset_index()
        m_summary['% Chốt Nóng'] = (m_summary['Chốt_Nóng'] / m_summary['Nhận_Tháng'] * 100).round(2)
        st.dataframe(m_summary, use_container_width=True, hide_index=True)

    with tab2:
        st.markdown("### Hiệu suất Chốt Nóng & Tổng Tỉ Lệ (%) từng Sale")
        st.dataframe(df_emp_final.sort_values(by=['THÁNG_NHẬN', 'Tổng Tỉ Lệ (%)'], ascending=[True, False]), 
                     use_container_width=True, hide_index=True)
else:
    st.info("Vui lòng tải file để em xuất báo cáo SF chuẩn nhất cho anh.")
