import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Sales Manager", layout="wide")

# CSS Dashboard chuyên nghiệp, tối giản, lề lối thoáng đạt
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
        end = datetime(int(row['Năm Nhận File']), int(row['Tháng nhận file']), 1)
        if pd.isna(row['THÁNG NHẬN LEAD']) or pd.isna(row['NĂM NHẬN LEAD']): return None
        start = datetime(int(row['NĂM NHẬN LEAD']), int(row['THÁNG NHẬN LEAD']), 1)
        return (end.year - start.year) * 12 + (end.month - start.month)
    except: return None

uploaded_file = st.sidebar.file_uploader("Upload Data File (2 Sheets)", type=["xlsx"])

if uploaded_file:
    # Đọc dữ liệu và ép kiểu LEAD ID ngay để tránh lỗi log
    df_sales = pd.read_excel(uploaded_file, sheet_name=0)
    df_leads = pd.read_excel(uploaded_file, sheet_name=1)
    df_sales['LEAD ID'] = df_sales['LEAD ID'].astype(str)
    df_leads['LEAD ID'] = df_leads['LEAD ID'].astype(str)

    # --- LOGIC XỬ LÝ DOANH SỐ & TỐC ĐỘ ---
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

    # --- LOGIC CHUYỂN ĐỔI (DỰA TRÊN CỘT TEAM Ở CẢ 2 SHEET) ---
    df_sf_only = df_sales[df_sales['SOURCE'] == 'SF']
    closed_ids = df_sf_only['LEAD ID'].unique()
    df_leads['Is_Closed'] = df_leads['LEAD ID'].isin(closed_ids)

    # 1. Chuyển đổi theo Team
    team_recv = df_leads.groupby('TEAM')['LEAD ID'].nunique().reset_index(name='Nhận')
    team_cls = df_leads[df_leads['Is_Closed']].groupby('TEAM')['LEAD ID'].nunique().reset_index(name='Chốt')
    team_sales = df_sales.groupby('TEAM')['DOANH SỐ THỰC'].sum().reset_index(name='Doanh số')
    df_team_final = pd.merge(team_recv, team_cls, on='TEAM', how='left').fillna(0)
    df_team_final = pd.merge(df_team_final, team_sales, on='TEAM', how='left').fillna(0)
    df_team_final['% Chuyển đổi'] = (df_team_final['Chốt'] / df_team_final['Nhận'] * 100).round(2)

    # 2. Chuyển đổi theo Nhân viên (Tháng + Tổng)
    df_leads['Tháng'] = df_leads['DATE ADDED'].dt.month
    l_monthly = df_leads.groupby(['OWNER', 'TEAM', 'Tháng']).size().reset_index(name='Nhận')
    c_monthly = df_leads[df_leads['Is_Closed']].groupby(['OWNER', 'Tháng']).size().reset_index(name='Chốt')
    
    perf_emp = pd.merge(l_monthly, c_monthly, on=['OWNER', 'Tháng'], how='left').fillna(0)
    # Tính tổng tích lũy từng nhân viên
    l_total = df_leads.groupby('OWNER').size().reset_index(name='Tổng Nhận')
    c_total = df_leads[df_leads['Is_Closed']].groupby('OWNER').size().reset_index(name='Tổng Chốt')
    perf_total = pd.merge(l_total, c_total, on='OWNER', how='left').fillna(0)
    perf_total['Tổng % Chuyển đổi'] = (perf_total['Tổng Chốt'] / perf_total['Tổng Nhận'] * 100).round(2)
    
    perf_display = pd.merge(perf_emp, perf_total[['OWNER', 'Tổng % Chuyển đổi']], on='OWNER', how='left')

    # 3. Sản phẩm & Thời gian
    prod_summary = df_sales.groupby('PRODUCT').size().reset_index(name='Số lượng chốt')
    time_summary = df_sales[df_sales['SOURCE']=='SF'].groupby('THỜI GIAN CHỐT').size().reset_index(name='Số lượng')

    # --- XUẤT FILE EXCEL QUẢN TRỊ (BÁO CÁO CHỐT) ---
    st.sidebar.markdown("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        # Formats
        h_fmt = workbook.add_format({'bold': True, 'bg_color': '#0f172a', 'font_color': 'white', 'border': 1, 'align': 'center'})
        red_fmt = workbook.add_format({'font_color': '#ef4444', 'bold': True, 'border': 1})
        std_fmt = workbook.add_format({'border': 1})

        # Sheet 1: DASHBOARD (Tổng hợp & Biểu đồ)
        df_team_final.to_excel(writer, sheet_name='DASHBOARD', index=False, startrow=2)
        ws_dash = writer.sheets['DASHBOARD']
        ws_dash.write('A1', f'BÁO CÁO QUẢN TRỊ - CHUYỂN ĐỔI CHUNG: {round(df_team_final["Chốt"].sum()/df_team_final["Nhận"].sum()*100, 2)}%', workbook.add_format({'bold': True, 'font_size': 14}))
        for i, col in enumerate(df_team_final.columns): ws_dash.write(2, i, col, h_fmt)
        
        c_sales = workbook.add_chart({'type': 'column'})
        c_sales.add_series({'categories': '=DASHBOARD!$A$4:$A$6', 'values': '=DASHBOARD!$D$4:$D$6', 'name': 'Doanh số Team', 'fill': {'color': '#1e293b'}})
        ws_dash.insert_chart('F2', c_sales)

        # Sheet 2: NHÂN VIÊN (Tổng tỉ lệ + Chi tiết tháng)
        perf_display.to_excel(writer, sheet_name='NHAN_VIEN', index=False)
        ws_emp = writer.sheets['NHAN_VIEN']
        ws_emp.set_column('A:G', 18)
        for i, col in enumerate(perf_display.columns): ws_emp.write(0, i, col, h_fmt)

        # Sheet 3: SẢN PHẨM & TỐC ĐỘ (Biểu đồ tròn)
        prod_summary.to_excel(writer, sheet_name='SAN_PHAM', index=False)
        ws_p = writer.sheets['SAN_PHAM']
        c_pie = workbook.add_chart({'type': 'pie'})
        c_pie.add_series({'categories': '=SAN_PHAM!$A$2:$A$10', 'values': '=SAN_PHAM!$B$2:$B$10', 'name': 'Cơ cấu Sản phẩm'})
        ws_p.insert_chart('D2', c_pie)

        # Sheet 4: CHI TIẾT DỮ LIỆU (Tô màu)
        df_sales.to_excel(writer, sheet_name='CHI_TIET_DATA', index=False)
        ws_det = writer.sheets['CHI_TIET_DATA']
        eval_idx = df_sales.columns.get_loc('ĐÁNH GIÁ')
        for r, v in enumerate(df_sales['ĐÁNH GIÁ']):
            if "Chậm" in str(v): ws_det.write(r+1, eval_idx, v, red_fmt)

    st.sidebar.download_button("DOWNLOAD FINAL REPORT", data=output.getvalue(), file_name=f"Manager_Report_{datetime.now().strftime('%d%m%Y')}.xlsx")

    # --- UI DASHBOARD WEB ---
    tab1, tab2, tab3 = st.tabs(["📊 TỔNG QUAN", "👥 HIỆU SUẤT SALE", "🔍 DỮ LIỆU CHI TIẾT"])
    
    with tab1:
        m1, m2, m3 = st.columns(3)
        m1.metric("TỔNG DOANH SỐ", f"${df_sales['DOANH SỐ THỰC'].sum():,.2f}")
        m2.metric("TỔNG LEAD SF NHẬN", f"{df_leads['LEAD ID'].nunique():,}")
        m3.metric("CHUYỂN ĐỔI CHUNG", f"{round(df_team_final['Chốt'].sum()/df_team_final['Nhận'].sum()*100, 2)}%")
        
        st.markdown("### Hiệu suất 3 Team (Dựa trên Lead thực nhận)")
        st.dataframe(df_team_final, use_container_width=True, hide_index=True)
        
        c1, c2 = st.columns(2, gap="large")
        with c1:
            st.markdown("### Cơ cấu Sản phẩm")
            st.dataframe(prod_summary.sort_values(by='Số lượng chốt', ascending=False), use_container_width=True, hide_index=True)
        with c2:
            st.markdown("### Tỉ lệ chốt theo thời gian")
            st.dataframe(time_summary, use_container_width=True, hide_index=True)

    with tab2:
        st.markdown("### Chi tiết Chuyển đổi Nhân viên (Tháng & Tổng tích lũy)")
        st.dataframe(perf_display.sort_values(by=['Tháng', 'Tổng % Chuyển đổi'], ascending=[True, False]), use_container_width=True, hide_index=True)

    with tab3:
        st.markdown("### Toàn bộ dữ liệu xử lý")
        st.dataframe(df_sales[['OWNER', 'TEAM', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC', 'ĐÁNH GIÁ']], use_container_width=True, hide_index=True)
else:
    st.info("Anh hãy tải file Excel (đã thêm cột TEAM ở Sheet 2) để bắt đầu phân tích.")
