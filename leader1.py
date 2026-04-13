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

uploaded_file = st.sidebar.file_uploader("Upload Data File", type=["xlsx"])

if uploaded_file:
    df_sales = pd.read_excel(uploaded_file, sheet_name=0)
    df_leads = pd.read_excel(uploaded_file, sheet_name=1)

    # Xử lý Logic Doanh số & Tốc độ
    df_sales['ANNUAL PREMIUM'] = df_sales['ANNUAL PREMIUM'].apply(clean_currency)
    df_sales['TARGET PREMIUM'] = df_sales['TARGET PREMIUM'].apply(clean_currency)
    df_sales['DOANH SỐ THỰC'] = df_sales[['ANNUAL PREMIUM', 'TARGET PREMIUM']].min(axis=1)
    df_sales['MONTHS_DIFF'] = df_sales.apply(calculate_diff, axis=1)
    df_sales['ĐÁNH GIÁ'] = df_sales['MONTHS_DIFF'].apply(lambda x: "Tự khai thác" if pd.isna(x) else (f"{int(x)} tháng - {'Chậm' if x > 6 else 'Nhanh'}"))

    # Tính toán Chuyển đổi CHI TIẾT NHÂN VIÊN (SF)
    df_sf = df_sales[df_sales['SOURCE'] == 'SF']
    l_cnt = df_leads.groupby('OWNER')['LEAD ID'].count().reset_index(name='Nhận')
    c_ids = df_sf['LEAD ID'].unique()
    c_cnt = df_leads[df_leads['LEAD ID'].isin(c_ids)].groupby('OWNER')['LEAD ID'].count().reset_index(name='Chốt')
    
    df_owner_conv = pd.merge(l_cnt, c_cnt, on='OWNER', how='left').fillna(0)
    df_owner_conv['% Chuyển đổi'] = (df_owner_conv['Chốt'] / df_owner_conv['Nhận'] * 100).round(2)
    
    # Tính Tổng chung
    total_received = df_owner_conv['Nhận'].sum()
    total_closed = df_owner_conv['Chốt'].sum()
    overall_rate = (total_closed / total_received * 100).round(2) if total_received > 0 else 0

    # Sidebar: Nút Export
    st.sidebar.markdown("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_owner_conv.to_excel(writer, sheet_name='SUMMARY_OWNERS', index=False)
        df_sales.to_excel(writer, sheet_name='DETAILS_ALL', index=False)
        
        workbook = writer.book
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#0f172a', 'font_color': 'white', 'border': 1})
        ws_sum = writer.sheets['SUMMARY_OWNERS']
        for col_num, value in enumerate(df_owner_conv.columns.values):
            ws_sum.write(0, col_num, value, header_fmt)
            ws_sum.set_column(col_num, col_num, 20)

    st.sidebar.download_button("Download Final Report", data=output.getvalue(), file_name="Manager_Report.xlsx")

    tab1, tab2, tab3 = st.tabs(["TỔNG QUAN", "CHI TIẾT SF", "CHI TIẾT CC"])

    with tab1:
        # Khối con số tổng chung
        mc1, mc2, mc3 = st.columns(3)
        with mc1: st.metric("TỔNG LEAD NHẬN", f"{total_received:,}")
        with mc2: st.metric("TỔNG LEAD CHỐT", f"{total_closed:,}")
        with mc3: st.metric("TỈ LỆ CHUYỂN ĐỔI CHUNG", f"{overall_rate}%")

        st.markdown("### Hiệu suất Chuyển đổi từng Nhân viên")
        st.dataframe(df_owner_conv.sort_values(by='% Chuyển đổi', ascending=False), use_container_width=True, hide_index=True)

        c1, c2 = st.columns([3, 2], gap="large")
        with c1:
            st.markdown("### Sản phẩm Tổng (SF + CC)")
            p_total = df_sales.groupby('PRODUCT')['LEAD ID'].count().reset_index(name='Số lượng').sort_values(by='Số lượng', ascending=False)
            st.dataframe(p_total, use_container_width=True, hide_index=True)
        with c2:
            st.markdown("### Top Doanh số thực tế")
            st.dataframe(df_sales.groupby('OWNER')['DOANH SỐ THỰC'].sum().sort_values(ascending=False).reset_index(), use_container_width=True, hide_index=True,
                         column_config={"DOANH SỐ THỰC": st.column_config.NumberColumn(format="$%.2f")})

    with tab2:
        st.markdown("### Dữ liệu xử lý nguồn SF")
        st.dataframe(df_sf[['OWNER', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC', 'ĐÁNH GIÁ']], use_container_width=True, hide_index=True,
                     column_config={"DOANH SỐ THỰC": st.column_config.NumberColumn(format="$%.2f")})

    with tab3:
        df_cc = df_sales[df_sales['SOURCE'] == 'CC']
        st.markdown("### Dữ liệu xử lý nguồn CC")
        st.dataframe(df_cc[['OWNER', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC']], use_container_width=True, hide_index=True,
                     column_config={"DOANH SỐ THỰC": st.column_config.NumberColumn(format="$%.2f")})
else:
    st.info("Anh hãy tải file lên để bắt đầu soi hiệu suất nhân viên.")
