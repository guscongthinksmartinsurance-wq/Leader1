import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Sales Manager", layout="wide")

# CSS Dashboard cao cấp: Font Inter, Navy theme, Card layout
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

    # Chuẩn hóa tên cột
    df_sales.columns = df_sales.columns.str.strip().str.upper()
    df_leads.columns = df_leads.columns.str.strip().str.upper()

    # Ép kiểu dữ liệu
    df_sales['LEAD ID'] = df_sales['LEAD ID'].astype(str)
    df_leads['LEAD ID'] = df_leads['LEAD ID'].astype(str)

    # Xử lý Logic doanh số & Tốc độ
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
    df_sales['ĐÁNH GIÁ'] = df_sales['MONTHS_DIFF'].apply(get_bucket)

    # Tách nguồn dữ liệu
    df_sf_only = df_sales[df_sales['SOURCE'] == 'SF']
    df_cc_only = df_sales[df_sales['SOURCE'] == 'CC']
    closed_ids = df_sf_only['LEAD ID'].unique()
    df_leads['IS_CLOSED'] = df_leads['LEAD ID'].isin(closed_ids)

    # 1. Chuyển đổi Team (Chỉ SF)
    t_recv = df_leads.groupby('TEAM')['LEAD ID'].nunique().reset_index(name='Nhận')
    t_cls = df_leads[df_leads['IS_CLOSED']].groupby('TEAM')['LEAD ID'].nunique().reset_index(name='Chốt')
    t_sales_sf = df_sf_only.groupby('TEAM')['DOANH SỐ THỰC'].sum().reset_index(name='Doanh số SF')
    df_team_sf = pd.merge(t_recv, t_cls, on='TEAM', how='left').fillna(0)
    df_team_sf = pd.merge(df_team_sf, t_sales_sf, on='TEAM', how='left').fillna(0)
    df_team_sf['% Chuyển đổi'] = (df_team_sf['Chốt'] / df_team_sf['Nhận'] * 100).round(2)

    # 2. Hiệu suất Nhân viên (Chỉ SF)
    date_col = [c for c in df_leads.columns if 'DATE' in c][0]
    df_leads['THÁNG_N'] = df_leads[date_col].dt.month
    l_m = df_leads.groupby(['OWNER', 'TEAM', 'THÁNG_N']).size().reset_index(name='Nhận_T')
    c_m = df_leads[df_leads['IS_CLOSED']].groupby(['OWNER', 'THÁNG_N']).size().reset_index(name='Chốt_T')
    perf_emp = pd.merge(l_m, c_m, on=['OWNER', 'THÁNG_N'], how='left').fillna(0)
    
    l_tot = df_leads.groupby('OWNER').size().reset_index(name='Tổng_Nhận')
    c_tot = df_leads[df_leads['IS_CLOSED']].groupby('OWNER').size().reset_index(name='Tổng_Chốt')
    perf_tot = pd.merge(l_tot, c_tot, on='OWNER', how='left').fillna(0)
    perf_tot['Tổng Tỉ Lệ (%)'] = (perf_tot['Tổng_Chốt'] / perf_tot['Tổng_Nhận'] * 100).round(2)
    perf_display = pd.merge(perf_emp, perf_tot[['OWNER', 'Tổng Tỉ Lệ (%)']], on='OWNER', how='left')

    # --- XUẤT FILE EXCEL (CHỈ NGUỒN SF) ---
    st.sidebar.markdown("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        # Định dạng
        h_fmt = workbook.add_format({'bold': True, 'bg_color': '#0f172a', 'font_color': 'white', 'border': 1, 'align': 'center'})
        red_fmt = workbook.add_format({'font_color': '#ef4444', 'bold': True, 'border': 1})
        num_fmt = workbook.add_format({'num_format': '$#,##0', 'border': 1})

        # Sheet 1: Dashboard Funnel
        df_team_sf.to_excel(writer, sheet_name='DASHBOARD_SF', index=False, startrow=2)
        ws_dash = writer.sheets['DASHBOARD_SF']
        overall_sf = round(df_team_sf['Chốt'].sum()/df_team_sf['Nhận'].sum()*100, 2) if df_team_sf['Nhận'].sum() > 0 else 0
        ws_dash.write('A1', f'BÁO CÁO SF FUNNEL - TỈ LỆ CHUNG: {overall_sf}%', workbook.add_format({'bold': True, 'font_size': 14}))
        for i, col in enumerate(df_team_sf.columns): ws_dash.write(2, i, col, h_fmt)
        
        c_sf = workbook.add_chart({'type': 'column'})
        c_sf.add_series({'categories': '=DASHBOARD_SF!$A$4:$A$6', 'values': '=DASHBOARD_SF!$D$4:$D$6', 'name': 'Doanh số SF', 'fill': {'color': '#1e293b'}})
        ws_dash.insert_chart('F2', c_sf)

        # Sheet 2: Nhân viên (Chỉ SF)
        perf_display.to_excel(writer, sheet_name='NHAN_VIEN_SF', index=False)
        ws_emp = writer.sheets['NHAN_VIEN_SF']
        ws_emp.set_column('A:G', 18)
        for i, col in enumerate(perf_display.columns): ws_emp.write(0, i, col, h_fmt)

        # Sheet 3: Sản phẩm SF
        prod_sf = df_sf_only.groupby('PRODUCT').size().reset_index(name='SỐ_LƯỢNG')
        prod_sf.to_excel(writer, sheet_name='SAN_PHAM_SF', index=False)
        ws_p = writer.sheets['SAN_PHAM_SF']
        c_p = workbook.add_chart({'type': 'pie'})
        c_p.add_series({'categories': '=SAN_PHAM_SF!$A$2:$A$10', 'values': '=SAN_PHAM_SF!$B$2:$B$10', 'name': 'Cơ cấu SP SF'})
        ws_p.insert_chart('D2', c_p)

    st.sidebar.download_button("Download SF Report", data=output.getvalue(), file_name=f"SF_Manager_Report_{datetime.now().strftime('%d%m%Y')}.xlsx")

    # --- UI WEB APP (HIỂN THỊ CẢ SF VÀ CC) ---
    tab1, tab2, tab3 = st.tabs(["📊 TỔNG QUAN", "👥 HIỆU SUẤT SALE", "🔍 DỮ LIỆU CHI TIẾT"])
    
    with tab1:
        c1, c2, c3 = st.columns(3)
        c1.metric("TỔNG DOANH SỐ (SF+CC)", f"${(df_sf_only['DOANH SỐ THỰC'].sum() + df_cc_only['DOANH SỐ THỰC'].sum()):,.2f}")
        c2.metric("TỔNG NHẬN SF", f"{df_leads['LEAD ID'].nunique():,}")
        c3.metric("CHUYỂN ĐỔI SF CHUNG", f"{overall_sf}%")
        
        st.markdown("### Hiệu suất 3 Team (Chỉ nguồn SF)")
        st.dataframe(df_team_sf, use_container_width=True, hide_index=True)
        
        c_l, c_r = st.columns(2, gap="large")
        with c_l:
            st.markdown("### Sản phẩm Tổng (SF + CC)")
            p_total = df_sales.groupby('PRODUCT').size().reset_index(name='SỐ LƯỢNG CHỐT')
            st.dataframe(p_total.sort_values(by='SỐ LƯỢNG CHỐT', ascending=False), use_container_width=True, hide_index=True)
        with c_r:
            st.markdown("### Doanh số tự thân (CC) theo Team")
            st.dataframe(df_cc_only.groupby('TEAM')['DOANH SỐ THỰC'].sum().reset_index(), use_container_width=True, hide_index=True,
                         column_config={"DOANH SỐ THỰC": st.column_config.NumberColumn(format="$%.2f")})

    with tab2:
        st.markdown("### Chuyển đổi Nhân viên (Chỉ tính trên nguồn SF)")
        st.dataframe(perf_display.sort_values(by=['THÁNG_N', 'Tổng Tỉ Lệ (%)'], ascending=[True, False]), use_container_width=True, hide_index=True)

    with tab3:
        st.markdown("### Toàn bộ dữ liệu xử lý")
        st.dataframe(df_sales[['OWNER', 'TEAM', 'LEAD ID', 'PRODUCT', 'SOURCE', 'DOANH SỐ THỰC', 'ĐÁNH GIÁ']], use_container_width=True, hide_index=True,
                     column_config={"DOANH SỐ THỰC": st.column_config.NumberColumn(format="$%.2f")})
else:
    st.info("Anh Công hãy tải file Excel có cột TEAM ở Sheet 2 để em bắt đầu nhé.")
