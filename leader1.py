import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Sales Manager", layout="wide")

# CSS Dashboard chuyên nghiệp, lề lối thoáng, tone màu Navy sang trọng
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

uploaded_file = st.sidebar.file_uploader("Upload Data File", type=["xlsx"])

if uploaded_file:
    # Đọc và chuẩn hóa (Chỉ lấy SF)
    df_sales_raw = pd.read_excel(uploaded_file, sheet_name=0)
    df_leads_raw = pd.read_excel(uploaded_file, sheet_name=1)
    
    df_sales_raw.columns = df_sales_raw.columns.str.strip().str.upper()
    df_leads_raw.columns = df_leads_raw.columns.str.strip().str.upper()
    
    df_sales_raw['LEAD ID'] = df_sales_raw['LEAD ID'].astype(str)
    df_leads_raw['LEAD ID'] = df_leads_raw['LEAD ID'].astype(str)
    
    # FILTER SF ONLY
    df_sales = df_sales_raw[df_sales_raw['SOURCE'] == 'SF'].copy()
    df_leads = df_leads_raw.copy()

    df_sales['ANNUAL PREMIUM'] = df_sales['ANNUAL PREMIUM'].apply(clean_currency)
    df_sales['TARGET PREMIUM'] = df_sales['TARGET PREMIUM'].apply(clean_currency)
    df_sales['DOANH SỐ THỰC'] = df_sales[['ANNUAL PREMIUM', 'TARGET PREMIUM']].min(axis=1)

    # Logic Chốt Nóng (Nhận tháng nào chốt tháng đó)
    df_leads['THÁNG_NHẬN'] = df_leads['DATE ADDED'].dt.month
    df_hot = pd.merge(
        df_leads[['OWNER', 'TEAM', 'LEAD ID', 'THÁNG_NHẬN']], 
        df_sales[['LEAD ID', 'THÁNG NHẬN FILE', 'PRODUCT', 'DOANH SỐ THỰC']], 
        on='LEAD ID', how='inner'
    )
    df_hot_closed = df_hot[df_hot['THÁNG_NHẬN'] == df_hot['THÁNG NHẬN FILE']]

    # Bảng chi tiết từng tháng
    monthly_data = []
    for m in sorted(df_leads['THÁNG_NHẬN'].unique()):
        recv = df_leads[df_leads['THÁNG_NHẬN'] == m].shape[0]
        hot_cls = df_hot_closed[df_hot_closed['THÁNG_NHẬN'] == m].shape[0]
        monthly_data.append({
            'Tháng': f"Tháng {int(m)}",
            'Lead Nhận': recv,
            'Chốt Nóng': hot_cls,
            'Tỉ lệ (%)': round(hot_cls/recv*100, 2) if recv > 0 else 0
        })
    df_monthly = pd.DataFrame(monthly_data)

    # Tính Hàng Tổng
    t_recv = df_monthly['Lead Nhận'].sum()
    t_hot = df_monthly['Chốt Nóng'].sum()
    t_rate = round(t_hot/t_recv*100, 2) if t_recv > 0 else 0
    df_total = pd.DataFrame([{'Tháng': 'TỔNG CỘNG', 'Lead Nhận': t_recv, 'Chốt Nóng': t_hot, 'Tỉ lệ (%)': t_rate}])
    df_report_final = pd.concat([df_monthly, df_total], ignore_index=True)

    # Thống kê Sản phẩm SF
    df_prod_sf = df_sales.groupby('PRODUCT').size().reset_index(name='Số lượng chốt').sort_values(by='Số lượng chốt', ascending=False)

    # --- XUẤT FILE EXCEL BÁO CÁO (SF ONLY) ---
    st.sidebar.markdown("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        h_fmt = workbook.add_format({'bold': True, 'bg_color': '#0f172a', 'font_color': 'white', 'border': 1, 'align': 'center'})
        df_report_final.to_excel(writer, sheet_name='SF_REPORT', index=False, startrow=1)
        ws = writer.sheets['SF_REPORT']
        for col_num, value in enumerate(df_report_final.columns.values): ws.write(1, col_num, value, h_fmt)
        
        # Biểu đồ trực quan trong Excel
        chart = workbook.add_chart({'type': 'column'})
        chart.add_series({
            'name': 'Chốt Nóng',
            'categories': f'=SF_REPORT!$A$3:$A${len(df_monthly)+2}',
            'values': f'=SF_REPORT!$C$3:$C${len(df_monthly)+2}',
            'fill': {'color': '#1e293b'}
        })
        ws.insert_chart('F2', chart)
        
        df_prod_sf.to_excel(writer, sheet_name='SAN_PHAM_SF', index=False)

    st.sidebar.download_button("Download SF Final Report", data=output.getvalue(), file_name="SF_Funnel_Report.xlsx")

    # --- UI WEB APP ---
    tab1, tab2 = st.tabs(["🚀 BÁO CÁO FUNNEL SF", "👥 CHI TIẾT NHÂN VIÊN"])
    
    with tab1:
        c1, c2, c3 = st.columns(3)
        c1.metric("TỔNG LEAD SF", f"{t_recv:,}")
        c2.metric("TỔNG CHỐT NÓNG", f"{t_hot:,}")
        c3.metric("TỈ LỆ CHUNG", f"{t_rate}%")
        
        st.markdown("### Chi tiết hiệu quả chốt theo tháng")
        st.dataframe(df_report_final, use_container_width=True, hide_index=True)
        
        st.markdown("### Biểu đồ so sánh Nhận vs Chốt Nóng")
        st.bar_chart(df_monthly.set_index('Tháng')[['Lead Nhận', 'Chốt Nóng']])

        st.markdown("### Cơ cấu Sản phẩm SF")
        st.dataframe(df_prod_sf, use_container_width=True, hide_index=True)

    with tab2:
        st.markdown("### Hiệu suất Chốt Nóng từng nhân viên")
        e_recv = df_leads.groupby(['OWNER', 'THÁNG_NHẬN']).size().reset_index(name='Nhận')
        e_hot = df_hot_closed.groupby(['OWNER', 'THÁNG_NHẬN']).size().reset_index(name='Chốt Nóng')
        df_e = pd.merge(e_recv, e_hot, on=['OWNER', 'THÁNG_NHẬN'], how='left').fillna(0)
        df_e['%'] = (df_e['Chốt Nóng'] / df_e['Nhận'] * 100).round(2)
        st.dataframe(df_e.sort_values(by=['THÁNG_NHẬN', '%'], ascending=[True, False]), use_container_width=True, hide_index=True)
else:
    st.info("Anh Công hãy tải file lên để em xử lý báo cáo SF chuẩn nhất nhé.")
