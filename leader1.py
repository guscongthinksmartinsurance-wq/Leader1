import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Sales Manager", layout="wide")

# CSS Dashboard chuyên nghiệp
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

uploaded_file = st.sidebar.file_uploader("Upload Data File (2 Sheets)", type=["xlsx"])

if uploaded_file:
    df_sales = pd.read_excel(uploaded_file, sheet_name=0)
    df_leads = pd.read_excel(uploaded_file, sheet_name=1)

    # --- XỬ LÝ LOGIC ---
    df_sales['ANNUAL PREMIUM'] = df_sales['ANNUAL PREMIUM'].apply(clean_currency)
    df_sales['TARGET PREMIUM'] = df_sales['TARGET PREMIUM'].apply(clean_currency)
    df_sales['DOANH SỐ THỰC'] = df_sales[['ANNUAL PREMIUM', 'TARGET PREMIUM']].min(axis=1)
    df_sales['MONTHS_DIFF'] = df_sales.apply(calculate_diff, axis=1)
    
    # Phân loại tốc độ chốt cho báo cáo
    def get_bucket(m):
        if pd.isna(m): return "Tự khai thác"
        if m <= 3: return "0-3 Tháng"
        if m <= 6: return "3-6 Tháng"
        if m <= 9: return "6-9 Tháng"
        if m <= 12: return "9-12 Tháng"
        return "> 12 Tháng"
    df_sales['PHÂN LOẠI THỜI GIAN'] = df_sales['MONTHS_DIFF'].apply(get_bucket)

    # Tính Chuyển đổi Nhân viên theo Tháng
    df_leads['Month'] = df_leads['DATE ADDED'].dt.month
    df_leads['Year'] = df_leads['DATE ADDED'].dt.year
    df_sales_sf = df_sales[df_sales['SOURCE'] == 'SF']
    
    # Nhận Lead theo tháng
    received_monthly = df_leads.groupby(['OWNER', 'Month']).size().reset_index(name='Nhận')
    # Chốt Lead theo tháng (khớp ID)
    closed_ids = df_sales_sf['LEAD ID'].unique()
    df_leads['Is_Closed'] = df_leads['LEAD ID'].isin(closed_ids)
    closed_monthly = df_leads[df_leads['Is_Closed']].groupby(['OWNER', 'Month']).size().reset_index(name='Chốt')
    
    perf_monthly = pd.merge(received_monthly, closed_monthly, on=['OWNER', 'Month'], how='left').fillna(0)
    perf_monthly['% Chuyển đổi'] = (perf_monthly['Chốt'] / perf_monthly['Nhận'] * 100).round(2)

    # Tổng hợp 3 Team
    team_metrics = []
    for team in ['G', 'H', 'T']:
        owners = df_sales[df_sales['TEAM'] == team]['OWNER'].unique()
        received = df_leads[df_leads['OWNER'].isin(owners)]['LEAD ID'].nunique()
        closed = df_sales[(df_sales['TEAM'] == team) & (df_sales['SOURCE'] == 'SF')]['LEAD ID'].nunique()
        team_metrics.append({
            'TEAM': team,
            'Nhận': received,
            'Chốt': closed,
            'Doanh số': df_sales[df_sales['TEAM'] == team]['DOANH SỐ THỰC'].sum()
        })
    df_team = pd.DataFrame(team_metrics)
    total_received = df_leads['LEAD ID'].nunique()
    total_closed = len(closed_ids)
    overall_conv = round(total_closed / total_received * 100, 2) if total_received > 0 else 0

    # --- XUẤT FILE EXCEL CÓ BIỂU ĐỒ & MÀU SẮC ---
    st.sidebar.markdown("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Sheet 1: DASHBOARD
        df_team.to_excel(writer, sheet_name='DASHBOARD', startrow=1, index=False)
        workbook = writer.book
        worksheet = writer.sheets['DASHBOARD']
        
        # Định dạng
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#0f172a', 'font_color': 'white', 'border': 1, 'align': 'center'})
        val_fmt = workbook.add_format({'border': 1, 'align': 'center'})
        money_fmt = workbook.add_format({'num_format': '$#,##0', 'border': 1})

        worksheet.write('A1', 'BÁO CÁO TỔNG HỢP 3 TEAM', workbook.add_format({'bold': True, 'font_size': 14}))
        for col_num, value in enumerate(df_team.columns.values):
            worksheet.write(1, col_num, value, header_fmt)
        
        # Vẽ biểu đồ cột doanh số
        chart_sales = workbook.add_chart({'type': 'column'})
        chart_sales.add_series({
            'name': 'Doanh số 3 Team',
            'categories': '=DASHBOARD!$A$3:$A$5',
            'values': '=DASHBOARD!$D$3:$D$5',
            'fill': {'color': '#1e293b'}
        })
        chart_sales.set_title({'name': 'Biểu đồ Doanh số theo Team'})
        worksheet.insert_chart('F2', chart_sales)

        # Sheet 2: HIỆU SUẤT NHÂN VIÊN
        perf_monthly.to_excel(writer, sheet_name='NHÂN VIÊN', index=False)
        ws_emp = writer.sheets['NHÂN VIÊN']
        ws_emp.set_column('A:E', 15)

        # Sheet 3: SẢN PHẨM & TỐC ĐỘ
        prod_data = df_sales.groupby('PRODUCT').size().reset_index(name='Số lượng')
        prod_data.to_excel(writer, sheet_name='SẢN PHẨM', index=False)
        ws_prod = writer.sheets['SẢN PHẨM']
        
        chart_pie = workbook.add_chart({'type': 'pie'})
        chart_pie.add_series({
            'categories': '=SẢN PHẨM!$A$2:$A$10',
            'values': '=SẢN PHẨM!$B$2:$B$10',
        })
        chart_pie.set_title({'name': 'Tỉ lệ Sản phẩm chốt'})
        ws_prod.insert_chart('D2', chart_pie)

    st.sidebar.download_button("EXPORT FINAL REPORT", data=output.getvalue(), file_name="Henry_Master_Report.xlsx")

    # --- GIAO DIỆN WEB ---
    tab1, tab2, tab3 = st.tabs(["📊 TỔNG QUAN", "👥 NHÂN VIÊN", "📦 SẢN PHẨM"])

    with tab1:
        c_m1, c_m2, c_m3 = st.columns(3)
        c_m1.metric("TỔNG DOANH SỐ", f"${df_sales['DOANH SỐ THỰC'].sum():,.2f}")
        c_m2.metric("TỔNG LEAD NHẬN", total_received)
        c_m3.metric("TỈ LỆ CHUYỂN ĐỔI CHUNG", f"{overall_rate}%")
        
        st.markdown("### Hiệu suất 3 Team")
        st.dataframe(df_team, use_container_width=True, hide_index=True)
        
        st.markdown("### Tỉ lệ chuyển đổi theo thời gian chốt")
        time_dist = df_sales['PHÂN LOẠI THỜI GIAN'].value_counts().reset_index()
        st.bar_chart(time_dist.set_index('PHÂN LOẠI THỜI GIAN'), color="#0f172a")

    with tab2:
        st.markdown("### Báo cáo chi tiết Nhân viên theo Tháng")
        st.dataframe(perf_monthly.sort_values(by=['Month', '% Chuyển đổi'], ascending=[True, False]), use_container_width=True, hide_index=True)

    with tab3:
        st.markdown("### Thống kê Sản phẩm được chốt")
        p_table = df_sales.groupby('PRODUCT')['DOANH SỐ THỰC'].agg(['count', 'sum']).reset_index()
        p_table.columns = ['Sản phẩm', 'Số lượng', 'Tổng doanh số ($)']
        st.dataframe(p_table.sort_values(by='Số lượng', ascending=False), use_container_width=True, hide_index=True)

else:
    st.info("Anh Công hãy upload file để em bắt đầu xuất báo cáo xịn cho anh.")
