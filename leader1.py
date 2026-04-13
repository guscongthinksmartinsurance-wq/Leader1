import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Sales Manager", layout="wide")

# CSS Dashboard cao cấp: Navy Theme, Font Inter sắc nét
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; background-color: #f8fafc; }
    h1 { font-weight: 800; color: #0f172a; letter-spacing: -0.05em; margin-bottom: 2rem; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px; }
    h3 { font-weight: 700; color: #1e293b; margin-top: 1.5rem; }
    [data-testid="stMetric"] { background-color: white; border: 1px solid #e2e8f0; padding: 25px !important; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
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

uploaded_file = st.sidebar.file_uploader("Upload Data File (SF Only)", type=["xlsx"])

if uploaded_file:
    # 1. Đọc và Chuẩn hóa (Chỉ SF)
    df_sales_raw = pd.read_excel(uploaded_file, sheet_name=0)
    df_leads_raw = pd.read_excel(uploaded_file, sheet_name=1)
    
    df_sales_raw.columns = df_sales_raw.columns.str.strip().str.upper()
    df_leads_raw.columns = df_leads_raw.columns.str.strip().str.upper()
    
    df_sales_raw['LEAD ID'] = df_sales_raw['LEAD ID'].astype(str)
    df_leads_raw['LEAD ID'] = df_leads_raw['LEAD ID'].astype(str)
    
    df_sales = df_sales_raw[df_sales_raw['SOURCE'] == 'SF'].copy()
    df_leads = df_leads_raw.copy()

    df_sales['ANNUAL PREMIUM'] = df_sales['ANNUAL PREMIUM'].apply(clean_currency)
    df_sales['TARGET PREMIUM'] = df_sales['TARGET PREMIUM'].apply(clean_currency)
    df_sales['DOANH SỐ THỰC'] = df_sales[['ANNUAL PREMIUM', 'TARGET PREMIUM']].min(axis=1)

    # 2. Logic Chốt Nóng (Nhận tháng nào chốt tháng đó)
    df_leads['THÁNG_NHẬN'] = df_leads['DATE ADDED'].dt.month
    df_hot = pd.merge(
        df_leads[['OWNER', 'TEAM', 'LEAD ID', 'THÁNG_NHẬN']], 
        df_sales[['LEAD ID', 'THÁNG NHẬN FILE', 'PRODUCT', 'DOANH SỐ THỰC']], 
        on='LEAD ID', how='inner'
    )
    df_hot_closed = df_hot[df_hot['THÁNG_NHẬN'] == df_hot['THÁNG NHẬN FILE']]

    # 3. Bảng Dashboard Funnel (DASHBOARD_SF)
    monthly_stats = []
    for m in sorted(df_leads['THÁNG_NHẬN'].unique()):
        recv = df_leads[df_leads['THÁNG_NHẬN'] == m].shape[0]
        hot_cls = df_hot_closed[df_hot_closed['THÁNG_NHẬN'] == m].shape[0]
        monthly_stats.append({
            'Tháng': f"Tháng {int(m)}",
            'Lead Nhận': recv,
            'Chốt Nóng': hot_cls,
            'Tỉ lệ Chốt Nóng (%)': round(hot_cls/recv*100, 2) if recv > 0 else 0
        })
    df_dash_monthly = pd.DataFrame(monthly_stats)
    t_recv, t_hot = df_dash_monthly['Lead Nhận'].sum(), df_dash_monthly['Chốt Nóng'].sum()
    t_rate = round(t_hot/t_recv*100, 2) if t_recv > 0 else 0
    df_dash_total = pd.concat([df_dash_monthly, pd.DataFrame([{'Tháng': 'TỔNG CỘNG', 'Lead Nhận': t_recv, 'Chốt Nóng': t_hot, 'Tỉ lệ Chốt Nóng (%)': t_rate}])], ignore_index=True)

    # 4. Bảng Nhân Viên (Chi tiết chốt nóng & Tích lũy)
    e_recv = df_leads.groupby(['OWNER', 'THÁNG_NHẬN']).size().reset_index(name='Lead Nhận')
    e_hot = df_hot_closed.groupby(['OWNER', 'THÁNG_NHẬN']).size().reset_index(name='Chốt Nóng')
    df_perf_emp = pd.merge(e_recv, e_hot, on=['OWNER', 'THÁNG_NHẬN'], how='left').fillna(0)
    df_perf_emp['% Chốt Nóng (Tháng)'] = (df_perf_emp['Chốt Nóng'] / df_perf_emp['Lead Nhận'] * 100).round(2)
    
    # Tính Tổng tích lũy 3 tháng của từng người
    l_tot = df_leads.groupby('OWNER').size().reset_index(name='Tổng Lead Nhận (3T)')
    c_hot_tot = df_hot_closed.groupby('OWNER').size().reset_index(name='Tổng Chốt Nóng (3T)')
    perf_overall = pd.merge(l_tot, c_hot_tot, on='OWNER', how='left').fillna(0)
    perf_overall['Tổng Tỉ Lệ Chốt Nóng (%)'] = (perf_overall['Tổng Chốt Nóng (3T)'] / perf_overall['Tổng Lead Nhận (3T)'] * 100).round(2)
    
    # Kết quả cuối cùng cho sheet Nhân Viên
    df_emp_final = pd.merge(df_perf_emp, perf_overall[['OWNER', 'Tổng Lead Nhận (3T)', 'Tổng Tỉ Lệ Chốt Nóng (%)']], on='OWNER', how='left')
    df_emp_final.rename(columns={'THÁNG_NHẬN': 'Tháng'}, inplace=True)

    # 5. Bảng Sản phẩm SF
    df_prod_sf = df_sales.groupby('PRODUCT').agg({'LEAD ID': 'count', 'DOANH SỐ THỰC': 'sum'}).reset_index()
    df_prod_sf.columns = ['Sản phẩm', 'Số lượng chốt', 'Doanh số ($)']
    df_prod_sf = df_prod_sf.sort_values(by='Số lượng chốt', ascending=False)

    # --- XUẤT FILE EXCEL 3 SHEETS ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        h_fmt = workbook.add_format({'bold': True, 'bg_color': '#0f172a', 'font_color': 'white', 'border': 1, 'align': 'center'})
        
        # SHEET 1: DASHBOARD
        df_dash_total.to_excel(writer, sheet_name='DASHBOARD_SF', index=False, startrow=2)
        ws1 = writer.sheets['DASHBOARD_SF']
        ws1.write('A1', f'HIỆU QUẢ FUNNEL - TỔNG TỈ LỆ: {t_rate}%', workbook.add_format({'bold': True, 'font_size': 14}))
        for i, col in enumerate(df_dash_total.columns): ws1.write(2, i, col, h_fmt)
        ws1.set_column('A:D', 20)
        chart1 = workbook.add_chart({'type': 'column'})
        chart1.add_series({'name': 'Chốt Nóng', 'categories': f'=DASHBOARD_SF!$A$4:$A${len(df_dash_monthly)+3}', 'values': f'=DASHBOARD_SF!$C$4:$C${len(df_dash_monthly)+3}', 'fill': {'color': '#0f172a'}})
        ws1.insert_chart('F2', chart1)

        # SHEET 2: NHAN_VIEN_SF
        df_emp_final.to_excel(writer, sheet_name='NHAN_VIEN_SF', index=False)
        ws2 = writer.sheets['NHAN_VIEN_SF']
        ws2.set_column('A:H', 22)
        for i, col in enumerate(df_emp_final.columns): ws2.write(0, i, col, h_fmt)

        # SHEET 3: SAN_PHAM_SF
        df_prod_sf.to_excel(writer, sheet_name='SAN_PHAM_SF', index=False)
        ws3 = writer.sheets['SAN_PHAM_SF']
        ws3.set_column('A:C', 20)
        for i, col in enumerate(df_prod_sf.columns): ws3.write(0, i, col, h_fmt)
        chart3 = workbook.add_chart({'type': 'pie'})
        chart3.add_series({'name': 'Cơ cấu SP', 'categories': f'=SAN_PHAM_SF!$A$2:$A${len(df_prod_sf)+1}', 'values': f'=SAN_PHAM_SF!$B$2:$B${len(df_prod_sf)+1}'})
        ws3.insert_chart('E2', chart3)

    st.sidebar.download_button("Download Full SF Report", data=output.getvalue(), file_name="Henry_SF_Manager_Report.xlsx")

    # --- UI WEB ---
    t_web1, t_web2 = st.tabs(["🚀 BÁO CÁO TỔNG HỢP", "👥 CHI TIẾT NHÂN VIÊN"])
    with t_web1:
        c_m1, c_m2, c_m3 = st.columns(3)
        c_m1.metric("TỔNG NHẬN SF", f"{t_recv:,}")
        c_m2.metric("TỔNG CHỐT NÓNG", f"{t_hot:,}")
        c_m3.metric("TỈ LỆ CHUNG", f"{t_rate}%")
        st.markdown("### Hiệu quả chốt nóng từng tháng")
        st.dataframe(df_dash_total, use_container_width=True, hide_index=True)
        st.bar_chart(df_dash_monthly.set_index('Tháng')[['Lead Nhận', 'Chốt Nóng']])
        st.markdown("### Sản phẩm SF (Số lượng & Doanh số)")
        st.dataframe(df_prod_sf, use_container_width=True, hide_index=True)

    with t_web2:
        st.markdown("### Hiệu suất Sale: Chốt nóng tháng & Tích lũy 3 tháng")
        st.dataframe(df_emp_final.sort_values(by=['Tháng', 'Tổng Tỉ Lệ Chốt Nóng (%)'], ascending=[True, False]), use_container_width=True, hide_index=True)
else:
    st.info("Anh Công hãy tải file để em xuất báo cáo SF chuẩn nhất nhé.")
