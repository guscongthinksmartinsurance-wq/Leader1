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

uploaded_file = st.sidebar.file_uploader("Upload Data File", type=["xlsx"])

if uploaded_file:
    df_sales = pd.read_excel(uploaded_file, sheet_name=0)
    df_leads = pd.read_excel(uploaded_file, sheet_name=1)

    df_sales['LEAD ID'] = df_sales['LEAD ID'].astype(str)
    df_leads['LEAD ID'] = df_leads['LEAD ID'].astype(str)

    df_sales['ANNUAL PREMIUM'] = df_sales['ANNUAL PREMIUM'].apply(clean_currency)
    df_sales['TARGET PREMIUM'] = df_sales['TARGET PREMIUM'].apply(clean_currency)
    df_sales['DOANH SỐ THỰC'] = df_sales[['ANNUAL PREMIUM', 'TARGET PREMIUM']].min(axis=1)
    df_sales['MONTHS_DIFF'] = df_sales.apply(calculate_diff, axis=1)
    
    def eval_speed(x):
        if pd.isna(x): return "Tự khai thác"
        return f"{int(x)} tháng - {'Chậm' if x > 6 else 'Nhanh'}"
    df_sales['ĐÁNH GIÁ'] = df_sales['MONTHS_DIFF'].apply(eval_speed)

    # Tính chuyển đổi từng nhân viên (Tổng + Theo tháng)
    df_sf_only = df_sales[df_sales['SOURCE'] == 'SF']
    closed_lead_ids = df_sf_only['LEAD ID'].unique()
    
    # 1. Tổng hợp theo tháng
    df_leads['Tháng'] = df_leads['DATE ADDED'].dt.month
    l_monthly = df_leads.groupby(['OWNER', 'Tháng']).size().reset_index(name='Nhận')
    c_monthly = df_leads[df_leads['LEAD ID'].isin(closed_lead_ids)].groupby(['OWNER', 'Tháng']).size().reset_index(name='Chốt')
    perf_monthly = pd.merge(l_monthly, c_monthly, on=['OWNER', 'Tháng'], how='left').fillna(0)
    
    # 2. Tổng hợp tích lũy của từng Sale
    l_total = df_leads.groupby('OWNER').size().reset_index(name='Tổng Nhận')
    c_total = df_leads[df_leads['LEAD ID'].isin(closed_lead_ids)].groupby('OWNER').size().reset_index(name='Tổng Chốt')
    perf_total = pd.merge(l_total, c_total, on='OWNER', how='left').fillna(0)
    perf_total['Tổng Tỉ Lệ (%)'] = (perf_total['Tổng Chốt'] / perf_total['Tổng Nhận'] * 100).round(2)
    
    # Ghép cột Tổng Tỉ Lệ vào bảng hàng tháng để anh Công dễ soi
    perf_final = pd.merge(perf_monthly, perf_total[['OWNER', 'Tổng Tỉ Lệ (%)']], on='OWNER', how='left')
    perf_final['% Chuyển đổi tháng'] = (perf_final['Chốt'] / perf_final['Nhận'] * 100).round(2)

    # Thống kê Team
    team_data = []
    for team in ['G', 'H', 'T']:
        owners = df_sales[df_sales['TEAM'] == team]['OWNER'].unique()
        recv = df_leads[df_leads['OWNER'].isin(owners)]['LEAD ID'].nunique()
        cls = df_sales[(df_sales['TEAM'] == team) & (df_sales['SOURCE'] == 'SF')]['LEAD ID'].nunique()
        team_data.append({
            'TEAM': team, 'Nhận': recv, 'Chốt': cls,
            'Doanh số ($)': df_sales[df_sales['TEAM'] == team]['DOANH SỐ THỰC'].sum()
        })
    df_team = pd.DataFrame(team_data)
    total_received = df_leads['LEAD ID'].nunique()
    total_closed = len(closed_lead_ids)
    overall_rate = round(total_closed/total_received*100, 2) if total_received > 0 else 0

    # --- XUẤT FILE EXCEL NÂNG CAO ---
    st.sidebar.markdown("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        # Định dạng màu sắc chuẩn chuyên nghiệp
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#0f172a', 'font_color': 'white', 'border': 1, 'align': 'center'})
        bg_fmt = workbook.add_format({'bg_color': '#f8fafc', 'border': 1})
        money_fmt = workbook.add_format({'num_format': '$#,##0', 'border': 1})
        red_fmt = workbook.add_format({'font_color': '#ef4444', 'bold': True, 'border': 1})
        pct_fmt = workbook.add_format({'num_format': '0.00"%"', 'border': 1, 'bold': True, 'bg_color': '#f1f5f9'})

        # Sheet DASHBOARD
        df_team.to_excel(writer, sheet_name='DASHBOARD', index=False, startrow=2)
        ws_dash = writer.sheets['DASHBOARD']
        ws_dash.write('A1', f'BÁO CÁO TỔNG QUAN - TỈ LỆ CHUNG: {overall_rate}%', workbook.add_format({'bold': True, 'font_size': 14}))
        for i, col in enumerate(df_team.columns): ws_dash.write(2, i, col, header_fmt)
        
        c_sales = workbook.add_chart({'type': 'column'})
        c_sales.add_series({'categories': '=DASHBOARD!$A$4:$A$6', 'values': '=DASHBOARD!$D$4:$D$6', 'name': 'Doanh số Team', 'fill': {'color': '#1e293b'}})
        ws_dash.insert_chart('F2', c_sales)

        # Sheet NHÂN VIÊN (Thêm cột Tổng Tỉ Lệ)
        perf_final.to_excel(writer, sheet_name='NHAN_VIEN', index=False)
        ws_emp = writer.sheets['NHAN_VIEN']
        ws_emp.set_column('A:G', 18)
        for i, col in enumerate(perf_final.columns): ws_emp.write(0, i, col, header_fmt)

        # Sheet CHI TIẾT (Tô màu Chậm)
        df_sales.to_excel(writer, sheet_name='CHI_TIET', index=False)
        ws_det = writer.sheets['CHI_TIET']
        ws_det.set_column('A:P', 15)
        eval_idx = df_sales.columns.get_loc('ĐÁNH GIÁ')
        for r, v in enumerate(df_sales['ĐÁNH GIÁ']):
            if "Chậm" in str(v): ws_det.write(r+1, eval_idx, v, red_fmt)
        
        # Sheet SẢN PHẨM (Biểu đồ tròn)
        p_data = df_sales.groupby('PRODUCT').size().reset_index(name='Qty')
        p_data.to_excel(writer, sheet_name='SAN_PHAM', index=False)
        ws_p = writer.sheets['SAN_PHAM']
        c_pie = workbook.add_chart({'type': 'pie'})
        c_pie.add_series({'categories': '=SAN_PHAM!$A$2:$A$10', 'values': '=SAN_PHAM!$B$2:$B$10', 'name': 'Tỉ lệ Sản phẩm'})
        ws_p.insert_chart('D2', c_pie)

    st.sidebar.download_button("Download Final Report", data=output.getvalue(), file_name="Henry_Manager_Final_Report.xlsx")

    # --- UI WEB APP ---
    t1, t2, t3 = st.tabs(["📊 TỔNG QUAN", "👥 HIỆU SUẤT SALE", "🔍 DỮ LIỆU CHI TIẾT"])
    
    with t1:
        c1, c2, c3 = st.columns(3)
        c1.metric("TỔNG DOANH SỐ", f"${df_sales['DOANH SỐ THỰC'].sum():,.2f}")
        c2.metric("TỔNG LEAD NHẬN", f"{total_received:,}")
        c3.metric("CHUYỂN ĐỔI CHUNG", f"{overall_rate}%")
        
        st.markdown("### Doanh số và Chuyển đổi theo Team")
        st.dataframe(df_team, use_container_width=True, hide_index=True)
        
        st.markdown("### Thống kê Sản phẩm chủ lực")
        st.dataframe(p_data.sort_values(by='Qty', ascending=False), use_container_width=True, hide_index=True)

    with t2:
        st.markdown("### Chi tiết Chuyển đổi Nhân viên")
        st.dataframe(perf_final.sort_values(by=['Tháng', 'Tổng Tỉ Lệ (%)'], ascending=[True, False]), 
                     use_container_width=True, hide_index=True,
                     column_config={"Tổng Tỉ Lệ (%)": st.column_config.NumberColumn(format="%.2f%%")})

    with t3:
        st.markdown("### Toàn bộ lịch sử chốt khách")
        st.dataframe(df_sales[['OWNER', 'TEAM', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC', 'ĐÁNH GIÁ']], 
                     use_container_width=True, hide_index=True,
                     column_config={"DOANH SỐ THỰC": st.column_config.NumberColumn(format="$%.2f")})
else:
    st.info("Anh hãy upload file để em bắt đầu xuất báo cáo chuyên nghiệp cho anh.")
