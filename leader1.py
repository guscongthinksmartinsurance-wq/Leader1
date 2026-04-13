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
    h1 { font-weight: 800; color: #0f172a; letter-spacing: -0.05em; margin-bottom: 2rem; border-bottom: 2px solid #e2e8f0; padding-bottom: 15px; }
    h3 { font-weight: 700; color: #1e293b; margin-top: 1.5rem; margin-bottom: 1rem; }
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

uploaded_file = st.sidebar.file_uploader("Upload Data File (2 Sheets)", type=["xlsx"])

if uploaded_file:
    df_sales = pd.read_excel(uploaded_file, sheet_name=0)
    df_leads = pd.read_excel(uploaded_file, sheet_name=1)

    df_sales.columns = df_sales.columns.str.strip().str.upper()
    df_leads.columns = df_leads.columns.str.strip().str.upper()

    df_sales['LEAD ID'] = df_sales['LEAD ID'].astype(str)
    df_leads['LEAD ID'] = df_leads['LEAD ID'].astype(str)

    # Lọc duy nhất nguồn SF
    df_sales = df_sales[df_sales['SOURCE'] == 'SF']
    df_sales['ANNUAL PREMIUM'] = df_sales['ANNUAL PREMIUM'].apply(clean_currency)
    df_sales['TARGET PREMIUM'] = df_sales['TARGET PREMIUM'].apply(clean_currency)
    df_sales['DOANH SỐ THỰC'] = df_sales[['ANNUAL PREMIUM', 'TARGET PREMIUM']].min(axis=1)

    # --- LOGIC CHỐT NÓNG (SAME MONTH CONVERSION) ---
    df_leads['THÁNG_NHẬN'] = df_leads['DATE ADDED'].dt.month
    
    # Lead ID đã chốt
    closed_ids = df_sales['LEAD ID'].unique()
    
    # Tạo bảng đối chiếu chốt nóng
    # Điều kiện chốt nóng: Lead ID nằm trong file chốt VÀ Tháng nhận (Sheet 2) == Tháng nhận file (Sheet 1)
    df_hot = pd.merge(
        df_leads[['OWNER', 'TEAM', 'LEAD ID', 'THÁNG_NHẬN']], 
        df_sales[['LEAD ID', 'THÁNG NHẬN FILE']], 
        on='LEAD ID', how='inner'
    )
    df_hot_closed = df_hot[df_hot['THÁNG_NHẬN'] == df_hot['THÁNG NHẬN FILE']]
    
    # Thống kê theo tháng cho Team & Nhân viên
    monthly_stats = []
    for m in sorted(df_leads['THÁNG_NHẬN'].unique()):
        recv = df_leads[df_leads['THÁNG_NHẬN'] == m].shape[0]
        hot_cls = df_hot_closed[df_hot_closed['THÁNG_NHẬN'] == m].shape[0]
        monthly_stats.append({
            'Tháng': f"Tháng {m}",
            'Tổng Nhận': recv,
            'Chốt Nóng': hot_cls,
            'Tỉ lệ Chốt Nóng (%)': round(hot_cls/recv*100, 2) if recv > 0 else 0
        })
    df_monthly_report = pd.DataFrame(monthly_stats)

    # Thống kê chi tiết từng Nhân viên (Chốt nóng)
    emp_recv = df_leads.groupby(['OWNER', 'THÁNG_NHẬN']).size().reset_index(name='Nhận')
    emp_hot = df_hot_closed.groupby(['OWNER', 'THÁNG_NHẬN']).size().reset_index(name='Chốt Nóng')
    df_emp_perf = pd.merge(emp_recv, emp_hot, on=['OWNER', 'THÁNG_NHẬN'], how='left').fillna(0)
    df_emp_perf['% Chốt Nóng'] = (df_emp_perf['Chốt Nóng'] / df_emp_perf['Nhận'] * 100).round(2)

    # --- XUẤT FILE EXCEL (SF ONLY - CHỐT NÓNG) ---
    st.sidebar.markdown("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        h_fmt = workbook.add_format({'bold': True, 'bg_color': '#0f172a', 'font_color': 'white', 'border': 1, 'align': 'center'})
        ws_dash = workbook.add_worksheet('BÁO CÁO CHỐT NÓNG')
        
        # Dashboard Tổng
        df_monthly_report.to_excel(writer, sheet_name='BÁO CÁO CHỐT NÓNG', index=False, startrow=2)
        ws_dash.write('A1', 'BÁO CÁO HIỆU QUẢ CHỐT NÓNG (SF FUNNEL)', workbook.add_format({'bold': True, 'font_size': 14}))
        for i, col in enumerate(df_monthly_report.columns): ws_dash.write(2, i, col, h_fmt)
        
        # Biểu đồ chốt nóng
        chart = workbook.add_chart({'type': 'column'})
        chart.add_series({
            'name': 'Chốt Nóng',
            'categories': '=BÁO CÁO CHỐT NÓNG!$A$4:$A$10',
            'values': '=BÁO CÁO CHỐT NÓNG!$C$4:$C$10',
            'fill': {'color': '#0f172a'}
        })
        ws_dash.insert_chart('F2', chart)

        # Sheet chi tiết nhân viên
        df_emp_perf.to_excel(writer, sheet_name='HIỆU SUẤT NHÂN VIÊN', index=False)
        ws_emp = writer.sheets['HIỆU SUẤT NHÂN VIÊN']
        for i, col in enumerate(df_emp_perf.columns): ws_emp.write(0, i, col, h_fmt)

    st.sidebar.download_button("DOWNLOAD SF HOT REPORT", data=output.getvalue(), file_name=f"SF_Hot_Report_{datetime.now().strftime('%d%m%Y')}.xlsx")

    # --- UI WEB ---
    t1, t2 = st.tabs(["🚀 CHỐT NÓNG THEO THÁNG", "👥 CHI TIẾT NHÂN VIÊN"])
    
    with t1:
        st.markdown("### Bức tranh Chốt Nóng toàn Team")
        st.dataframe(df_monthly_report, use_container_width=True, hide_index=True)
        
        st.markdown("### Biểu đồ Tỉ lệ Chốt Nóng (%)")
        st.line_chart(df_monthly_report.set_index('Tháng')['Tỉ lệ Chốt Nóng (%)'], color="#0f172a")

    with t2:
        st.markdown("### Hiệu suất Chốt Nóng từng Nhân viên")
        st.dataframe(df_emp_perf.sort_values(by=['THÁNG_NHẬN', '% Chốt Nóng'], ascending=[True, False]), use_container_width=True, hide_index=True)
else:
    st.info("Anh Công hãy tải file lên để em soi tỉ lệ chốt nóng của anh em nhé.")
