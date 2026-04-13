import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Sales Manager", layout="wide")

# CSS Dashboard cao cấp, tối giản, chuyên nghiệp
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; background-color: #f8fafc; }
    .main { padding: 2rem 5rem; }
    h1 { font-weight: 800; color: #0f172a; letter-spacing: -0.05em; margin-bottom: 2rem; border-bottom: 2px solid #e2e8f0; padding-bottom: 15px; }
    h3 { font-weight: 700; color: #1e293b; margin-top: 1.5rem; margin-bottom: 1rem; }
    [data-testid="stMetric"] { background-color: white; border: 1px solid #e2e8f0; padding: 25px !important; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
    [data-testid="stMetricValue"] { color: #0f172a; font-weight: 700 !important; }
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

uploaded_file = st.sidebar.file_uploader("Upload Data File (2 Sheets)", type=["xlsx"])

if uploaded_file:
    # 1. Đọc và Chuẩn hóa (Lọc SF ngay lập tức)
    df_sales_raw = pd.read_excel(uploaded_file, sheet_name=0)
    df_leads_raw = pd.read_excel(uploaded_file, sheet_name=1)
    
    df_sales_raw.columns = df_sales_raw.columns.str.strip().str.upper()
    df_leads_raw.columns = df_leads_raw.columns.str.strip().str.upper()
    
    # Ép kiểu Lead ID để tránh lỗi log
    df_sales_raw['LEAD ID'] = df_sales_raw['LEAD ID'].astype(str)
    df_leads_raw['LEAD ID'] = df_leads_raw['LEAD ID'].astype(str)
    
    # CHỈ LẤY NGUỒN SF
    df_sales = df_sales_raw[df_sales_raw['SOURCE'] == 'SF'].copy()
    df_leads = df_leads_raw.copy()

    df_sales['ANNUAL PREMIUM'] = df_sales['ANNUAL PREMIUM'].apply(clean_currency)
    df_sales['TARGET PREMIUM'] = df_sales['TARGET PREMIUM'].apply(clean_currency)
    df_sales['DOANH SỐ THỰC'] = df_sales[['ANNUAL PREMIUM', 'TARGET PREMIUM']].min(axis=1)

    # 2. Logic Chốt Nóng (Same Month)
    df_leads['THÁNG_NHẬN'] = df_leads['DATE ADDED'].dt.month
    df_hot = pd.merge(
        df_leads[['OWNER', 'TEAM', 'LEAD ID', 'THÁNG_NHẬN']], 
        df_sales[['LEAD ID', 'THÁNG NHẬN FILE', 'PRODUCT', 'DOANH SỐ THỰC']], 
        on='LEAD ID', how='inner'
    )
    df_hot_closed = df_hot[df_hot['THÁNG_NHẬN'] == df_hot['THÁNG NHẬN FILE']]

    # 3. Tổng hợp Báo cáo Tháng (Có hàng Tổng)
    monthly_data = []
    months = sorted(df_leads['THÁNG_NHẬN'].unique())
    for m in months:
        recv = df_leads[df_leads['THÁNG_NHẬN'] == m].shape[0]
        hot_cls = df_hot_closed[df_hot_closed['THÁNG_NHẬN'] == m].shape[0]
        monthly_data.append({
            'Tháng': f"Tháng {int(m)}",
            'Lead Nhận': recv,
            'Chốt Nóng': hot_cls,
            'Tỉ lệ (%)': round(hot_cls/recv*100, 2) if recv > 0 else 0
        })
    
    df_monthly = pd.DataFrame(monthly_data)
    
    # Thêm hàng TỔNG CỘNG
    total_recv = df_monthly['Lead Nhận'].sum()
    total_hot = df_monthly['Chốt Nóng'].sum()
    total_rate = round(total_hot/total_recv*100, 2) if total_recv > 0 else 0
    
    total_row = pd.DataFrame([{'Tháng': 'TỔNG CỘNG', 'Lead Nhận': total_recv, 'Chốt Nóng': total_hot, 'Tỉ lệ (%)': total_rate}])
    df_monthly_with_total = pd.concat([df_monthly, total_row], ignore_index=True)

    # 4. Thống kê Sản phẩm SF
    df_prod_sf = df_sales.groupby('PRODUCT').agg({'LEAD ID': 'count', 'DOANH SỐ THỰC': 'sum'}).reset_index()
    df_prod_sf.columns = ['Sản phẩm', 'Số lượng chốt', 'Tổng doanh số ($)']
    df_prod_sf = df_prod_sf.sort_values(by='Số lượng chốt', ascending=False)

    # --- XUẤT FILE EXCEL QUẢN TRỊ (SF ONLY) ---
    st.sidebar.markdown("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        # Định dạng
        h_fmt = workbook.add_format({'bold': True, 'bg_color': '#0f172a', 'font_color': 'white', 'border': 1, 'align': 'center'})
        total_fmt = workbook.add_format({'bold': True, 'bg_color': '#f1f5f9', 'border': 1, 'align': 'center'})
        std_fmt = workbook.add_format({'border': 1, 'align': 'center'})
        
        # Sheet 1: BÁO CÁO SF TỔNG HỢP
        df_monthly_with_total.to_excel(writer, sheet_name='BAO_CAO_SF', index=False, startrow=2)
        ws_sf = writer.sheets['BAO_CAO_SF']
        ws_sf.write('A1', f'BÁO CÁO HIỆU QUẢ SF - CHUYỂN ĐỔI TỔNG: {total_rate}%', workbook.add_format({'bold': True, 'font_size': 14}))
        
        # Apply header format
        for col_num, value in enumerate(df_monthly_with_total.columns.values):
            ws_sf.write(2, col_num, value, h_fmt)
        
        # Highlight hàng Tổng trong Excel
        last_row = len(df_monthly_with_total) + 2
        for col_num in range(len(df_monthly_with_total.columns)):
            ws_sf.write(last_row, col_num, df_monthly_with_total.iloc[-1, col_num], total_fmt)
        
        # Sheet 2: SẢN PHẨM SF
        df_prod_sf.to_excel(writer, sheet_name='SAN_PHAM_SF', index=False)
        ws_p = writer.sheets['SAN_PHAM_SF']
        ws_p.set_column('A:C', 20)
        for i, col in enumerate(df_prod_sf.columns): ws_p.write(0, i, col, h_fmt)

    st.sidebar.download_button("DOWNLOAD SF FINAL REPORT", data=output.getvalue(), file_name=f"SF_Final_Report_{datetime.now().strftime('%d%m%Y')}.xlsx")

    # --- UI WEB APP ---
    tab1, tab2, tab3 = st.tabs(["📊 BÁO CÁO TỔNG", "👥 CHI TIẾT NHÂN VIÊN", "🔍 RAW DATA SF"])
    
    with tab1:
        c1, c2, c3 = st.columns(3)
        c1.metric("TỔNG LEAD SF", f"{total_recv:,}")
        c2.metric("TỔNG CHỐT NÓNG", f"{total_hot:,}")
        c3.metric("TỈ LỆ CHUYỂN ĐỔI", f"{total_rate}%")
        
        st.markdown("### Tỉ lệ chốt nóng theo tháng")
        st.dataframe(df_monthly_with_total, use_container_width=True, hide_index=True)
        
        st.markdown("### Thống kê Sản phẩm SF")
        st.dataframe(df_prod_sf, use_container_width=True, hide_index=True)

    with tab2:
        st.markdown("### Hiệu suất Chốt Nóng từng nhân viên")
        e_recv = df_leads.groupby(['OWNER', 'THÁNG_NHẬN']).size().reset_index(name='Nhận')
        e_hot = df_hot_closed.groupby(['OWNER', 'THÁNG_NHẬN']).size().reset_index(name='Chốt Nóng')
        df_e = pd.merge(e_recv, e_hot, on=['OWNER', 'THÁNG_NHẬN'], how='left').fillna(0)
        df_e['% Chốt Nóng'] = (df_e['Chốt Nóng'] / df_e['Nhận'] * 100).round(2)
        st.dataframe(df_e.sort_values(by=['THÁNG_NHẬN', '% Chốt Nóng'], ascending=[True, False]), use_container_width=True, hide_index=True)

    with tab3:
        st.markdown("### Dữ liệu nguồn SF đã xử lý")
        st.dataframe(df_sales[['OWNER', 'TEAM', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC']], use_container_width=True, hide_index=True)
else:
    st.info("Anh hãy tải file Excel lên để em xuất báo cáo SF chuẩn chỉnh cho anh.")
