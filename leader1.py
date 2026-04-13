import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Sales Manager", layout="wide")

# CSS chuyên nghiệp, lề lối thoáng đạt
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
    # Đọc dữ liệu
    df_sales = pd.read_excel(uploaded_file, sheet_name=0)
    df_leads = pd.read_excel(uploaded_file, sheet_name=1)

    # FIX LỖI PYARROW: Ép kiểu LEAD ID về String
    df_sales['LEAD ID'] = df_sales['LEAD ID'].astype(str)
    df_leads['LEAD ID'] = df_leads['LEAD ID'].astype(str)

    # --- LOGIC XỬ LÝ ---
    df_sales['ANNUAL PREMIUM'] = df_sales['ANNUAL PREMIUM'].apply(clean_currency)
    df_sales['TARGET PREMIUM'] = df_sales['TARGET PREMIUM'].apply(clean_currency)
    df_sales['DOANH SỐ THỰC'] = df_sales[['ANNUAL PREMIUM', 'TARGET PREMIUM']].min(axis=1)
    df_sales['MONTHS_DIFF'] = df_sales.apply(calculate_diff, axis=1)
    
    # Phân loại thời gian chốt
    def get_bucket(m):
        if pd.isna(m): return "Tự khai thác"
        if m <= 3: return "0-3 Tháng"
        if m <= 6: return "3-6 Tháng"
        if m <= 9: return "6-9 Tháng"
        if m <= 12: return "9-12 Tháng"
        return "> 12 Tháng"
    df_sales['SPEED'] = df_sales['MONTHS_DIFF'].apply(get_bucket)

    # Thống kê Nhân viên theo Tháng
    df_leads['Tháng'] = df_leads['DATE ADDED'].dt.month
    l_sum = df_leads.groupby(['OWNER', 'Tháng']).size().reset_index(name='Nhận')
    c_ids = df_sales[df_sales['SOURCE'] == 'SF']['LEAD ID'].unique()
    c_sum = df_leads[df_leads['LEAD ID'].isin(c_ids)].groupby(['OWNER', 'Tháng']).size().reset_index(name='Chốt')
    perf_emp = pd.merge(l_sum, c_sum, on=['OWNER', 'Tháng'], how='left').fillna(0)
    perf_emp['% Chuyển đổi'] = (perf_emp['Chốt'] / perf_emp['Nhận'] * 100).round(2)

    # Thống kê Team
    team_data = []
    for team in ['G', 'H', 'T']:
        owners = df_sales[df_sales['TEAM'] == team]['OWNER'].unique()
        recv = df_leads[df_leads['OWNER'].isin(owners)]['LEAD ID'].nunique()
        cls = df_sales[(df_sales['TEAM'] == team) & (df_sales['SOURCE'] == 'SF')]['LEAD ID'].nunique()
        team_data.append({
            'TEAM': team,
            'Nhận': recv,
            'Chốt': cls,
            'Doanh số ($)': df_sales[df_sales['TEAM'] == team]['DOANH SỐ THỰC'].sum()
        })
    df_team = pd.DataFrame(team_data)
    total_received = df_leads['LEAD ID'].nunique()
    total_closed = len(c_ids)
    overall_rate = round(total_closed/total_received*100, 2) if total_received > 0 else 0

    # --- XUẤT FILE EXCEL BÁO CÁO QUẢN TRỊ ---
    st.sidebar.markdown("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        h_fmt = workbook.add_format({'bold': True, 'bg_color': '#0f172a', 'font_color': 'white', 'border': 1, 'align': 'center'})
        red_fmt = workbook.add_format({'font_color': '#ef4444', 'bold': True})
        
        # Sheet Dashboard
        df_team.to_excel(writer, sheet_name='DASHBOARD', index=False, startrow=1)
        ws_dash = writer.sheets['DASHBOARD']
        ws_dash.write('A1', f'BÁO CÁO TỔNG: Tỉ lệ chuyển đổi chung {overall_rate}%', workbook.add_format({'bold': True, 'font_size': 12}))
        for i, col in enumerate(df_team.columns): ws_dash.write(1, i, col, h_fmt)
        
        # Biểu đồ Doanh số
        c_sales = workbook.add_chart({'type': 'column'})
        c_sales.add_series({'categories': '=DASHBOARD!$A$3:$A$5', 'values': '=DASHBOARD!$D$3:$D$5', 'name': 'Doanh số'})
        ws_dash.insert_chart('F2', c_sales)

        # Sheet Sản phẩm & Nhân viên
        p_data = df_sales.groupby('PRODUCT').size().reset_index(name='Qty')
        p_data.to_excel(writer, sheet_name='SAN_PHAM', index=False)
        ws_p = writer.sheets['SAN_PHAM']
        c_pie = workbook.add_chart({'type': 'pie'})
        c_pie.add_series({'categories': '=SAN_PHAM!$A$2:$A$10', 'values': '=SAN_PHAM!$B$2:$B$10'})
        ws_p.insert_chart('D2', c_pie)
        
        perf_emp.to_excel(writer, sheet_name='NHAN_VIEN', index=False)
        df_sales.to_excel(writer, sheet_name='CHI_TIET', index=False)
        ws_det = writer.sheets['CHI_TIET']
        speed_idx = df_sales.columns.get_loc('SPEED')
        for r, v in enumerate(df_sales['SPEED']):
            if "Chậm" in str(v) or "> 12" in str(v): ws_det.write(r+1, speed_idx, v, red_fmt)

    st.sidebar.download_button("DOWNLOAD FINAL REPORT", data=output.getvalue(), file_name="Henry_Manager_Report.xlsx")

    # --- HIỂN THỊ WEB ---
    t1, t2, t3 = st.tabs(["📊 TỔNG QUAN", "👥 NHÂN VIÊN", "🔍 CHI TIẾT"])
    
    with t1:
        c_m1, c_m2, c_m3 = st.columns(3)
        c_m1.metric("TỔNG DOANH SỐ", f"${df_sales['DOANH SỐ THỰC'].sum():,.2f}")
        c_m2.metric("TỔNG NHẬN", total_received)
        c_m3.metric("CHUYỂN ĐỔI CHUNG", f"{overall_rate}%")
        
        st.markdown("### Hiệu suất 3 Team")
        st.dataframe(df_team, width='stretch', hide_index=True)
        st.markdown("### Sản phẩm Tổng")
        st.dataframe(p_data.sort_values(by='Qty', ascending=False), width='stretch', hide_index=True)

    with t2:
        st.markdown("### Hiệu suất Nhân viên theo Tháng")
        st.dataframe(perf_emp.sort_values(by=['Tháng', '% Chuyển đổi'], ascending=[True, False]), width='stretch', hide_index=True)

    with t3:
        st.markdown("### Dữ liệu chốt chi tiết")
        st.dataframe(df_sales[['OWNER', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC', 'SPEED']], width='stretch', hide_index=True)
else:
    st.info("Anh hãy upload file để em bắt đầu soi số liệu nhé.")
