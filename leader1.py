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
    h3 { font-weight: 700; color: #1e293b; margin-top: 2rem; margin-bottom: 1rem; }
    [data-testid="stMetric"] { background-color: white; border: 1px solid #e2e8f0; padding: 25px !important; border-radius: 12px; }
    [data-testid="stMetricValue"] { color: #0f172a; font-weight: 700 !important; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; background-color: #f1f5f9; padding: 8px; border-radius: 12px; }
    .stTabs [data-baseweb="tab"] { height: 45px; border-radius: 8px; border: none; padding: 0 30px; font-weight: 600; color: #64748b; }
    .stTabs [aria-selected="true"] { background-color: #0f172a !important; color: white !important; }
    .stDownloadButton > button { background-color: #0f172a !important; color: white !important; border-radius: 8px !important; width: 100%; padding: 12px; font-weight: 700; border: none !important; }
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
        end = datetime(int(row['Năm Nhận File']), int(row['Tháng nhận file']), 1)
        if pd.isna(row['THÁNG NHẬN LEAD']) or pd.isna(row['NĂM NHẬN LEAD']): return None
        start = datetime(int(row['NĂM NHẬN LEAD']), int(row['THÁNG NHẬN LEAD']), 1)
        diff = (end.year - start.year) * 12 + (end.month - start.month)
        return diff
    except: return None

uploaded_file = st.sidebar.file_uploader("Upload Data File", type=["xlsx"])

if uploaded_file:
    df_sales = pd.read_excel(uploaded_file, sheet_name=0)
    df_leads = pd.read_excel(uploaded_file, sheet_name=1)

    df_sales['ANNUAL PREMIUM'] = df_sales['ANNUAL PREMIUM'].apply(clean_currency)
    df_sales['TARGET PREMIUM'] = df_sales['TARGET PREMIUM'].apply(clean_currency)
    df_sales['DOANH SỐ THỰC'] = df_sales[['ANNUAL PREMIUM', 'TARGET PREMIUM']].min(axis=1)
    
    df_sales['MONTHS_DIFF'] = df_sales.apply(calculate_diff, axis=1)
    df_sales['SPEED_EVAL'] = df_sales['MONTHS_DIFF'].apply(lambda x: "Tự khai thác" if pd.isna(x) else (f"{int(x)} tháng - Chậm" if x > 6 else f"{int(x)} tháng - Nhanh"))

    st.sidebar.markdown("---")
    export_out = io.BytesIO()
    with pd.ExcelWriter(export_out, engine='xlsxwriter') as writer:
        df_sales.to_excel(writer, index=False, sheet_name='Sales_Report')
    st.sidebar.download_button("EXPORT FINAL REPORT", data=export_out.getvalue(), file_name="Sales_Manager_Report.xlsx")

    tab1, tab2, tab3 = st.tabs(["TỔNG QUAN", "NGUỒN SF", "NGUỒN CC"])

    with tab1:
        st.metric("TỔNG DOANH SỐ TOÀN TEAM", f"${df_sales['DOANH SỐ THỰC'].sum():,.2f}")
        c1, c2 = st.columns([3, 2], gap="large")
        with c1:
            st.markdown("### Doanh số theo Team")
            t_data = df_sales.groupby('TEAM')['DOANH SỐ THỰC'].sum().reset_index()
            st.bar_chart(t_data.set_index('TEAM'), color="#0f172a")
        with c2:
            st.markdown("### Top Owners")
            o_data = df_sales.groupby('OWNER')['DOANH SỐ THỰC'].sum().sort_values(ascending=False).reset_index()
            st.dataframe(o_data, use_container_width=True, hide_index=True)

    with tab2:
        df_sf = df_sales[df_sales['SOURCE'] == 'SF']
        
        st.markdown("### Hiệu quả chuyển đổi SF")
        df_leads['Month_Year'] = df_leads['DATE ADDED'].dt.strftime('%m/%Y')
        l_summary = df_leads.groupby(['OWNER', 'Month_Year'])['LEAD ID'].count().reset_index(name='Nhận')
        
        closed_ids = df_sf['LEAD ID'].unique()
        df_leads['Is_Closed'] = df_leads['LEAD ID'].isin(closed_ids)
        c_summary = df_leads[df_leads['Is_Closed'] == True].groupby(['OWNER', 'Month_Year'])['LEAD ID'].count().reset_index(name='Chốt')
        
        conv_final = pd.merge(l_summary, c_summary, on=['OWNER', 'Month_Year'], how='left').fillna(0)
        conv_final['% Chuyển đổi'] = (conv_final['Chốt'] / conv_final['Nhận'] * 100).round(2)
        st.dataframe(conv_final.sort_values(by=['Month_Year', '% Chuyển đổi'], ascending=False), use_container_width=True, hide_index=True)

        cs1, cs2 = st.columns(2, gap="medium")
        with cs1:
            st.markdown("### Sản phẩm SF")
            p_sf = df_sf.groupby('PRODUCT')['LEAD ID'].count().reset_index(name='Số lượng').sort_values(by='Số lượng', ascending=False)
            st.dataframe(p_sf, use_container_width=True, hide_index=True)
        with cs2:
            st.markdown("### Thống kê Tốc độ")
            st.dataframe(df_sf['SPEED_EVAL'].value_counts().reset_index(name='Số lượng'), use_container_width=True, hide_index=True)

        st.markdown("### Chi tiết chốt SF")
        def style_eval(v):
            return 'color: #dc2626; font-weight: bold;' if 'Chậm' in str(v) else ''
        st.dataframe(df_sf[['OWNER', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC', 'SPEED_EVAL']].style.applymap(style_eval, subset=['SPEED_EVAL']), use_container_width=True, hide_index=True)

    with tab3:
        df_cc = df_sales[df_sales['SOURCE'] == 'CC']
        cc1, cc2 = st.columns([1, 2], gap="medium")
        with cc1:
            st.markdown("### Sản phẩm CC")
            p_cc = df_cc.groupby('PRODUCT')['LEAD ID'].count().reset_index(name='Số lượng').sort_values(by='Số lượng', ascending=False)
            st.dataframe(p_cc, use_container_width=True, hide_index=True)
        with cc2:
            st.markdown("### Chi tiết chốt CC")
            st.dataframe(df_cc[['OWNER', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC']], use_container_width=True, hide_index=True)
else:
    st.info("Vui lòng tải file để bắt đầu.")
