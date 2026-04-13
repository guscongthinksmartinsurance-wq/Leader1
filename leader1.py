import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Sales Manager", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        background-color: #f8fafc;
    }
    
    .main {
        padding: 2rem;
    }

    .stMetric {
        background-color: #ffffff;
        padding: 24px;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
    }

    div[data-testid="stMetricValue"] {
        font-size: 28px;
        font-weight: 700;
        color: #1e293b;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #f1f5f9;
        padding: 8px;
        border-radius: 12px;
    }

    .stTabs [data-baseweb="tab"] {
        height: 45px;
        border-radius: 8px;
        border: none;
        padding: 0 24px;
        background-color: transparent;
        color: #64748b;
        font-weight: 600;
    }

    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #0f172a !important;
        box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);
    }

    h1 {
        font-weight: 700;
        color: #0f172a;
        letter-spacing: -0.025em;
        margin-bottom: 2rem;
    }

    h3 {
        font-weight: 600;
        color: #334155;
        font-size: 1.1rem;
        margin-top: 1.5rem;
    }

    .dataframe {
        border-radius: 8px;
        overflow: hidden;
        border: 1px solid #e2e8f0;
    }

    .stButton>button {
        background-color: #0f172a;
        color: white;
        border-radius: 8px;
        font-weight: 600;
        border: none;
        padding: 0.6rem 1rem;
        transition: all 0.2s;
    }

    .stButton>button:hover {
        background-color: #334155;
        border: none;
    }

    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e2e8f0;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("Sales Manager")

def clean_currency(value):
    if isinstance(value, str):
        return float(value.replace('$', '').replace(',', '').strip())
    return value

def calculate_month_diff(row):
    try:
        end_date = datetime(int(row['Năm Nhận File']), int(row['Tháng nhận file']), 1)
        if pd.isna(row['THÁNG NHẬN LEAD']) or pd.isna(row['NĂM NHẬN LEAD']):
            return None
        start_date = datetime(int(row['NĂM NHẬN LEAD']), int(row['THÁNG NHẬN LEAD']), 1)
        return (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month)
    except:
        return None

def classify_speed(m):
    if pd.isna(m): return "Tự khai thác"
    return f"{int(m)} tháng - {'Chậm' if m > 6 else 'Nhanh'}"

uploaded_file = st.sidebar.file_uploader("Upload Data File", type=["xlsx"])

if uploaded_file:
    df_sales = pd.read_excel(uploaded_file, sheet_name=0)
    df_leads = pd.read_excel(uploaded_file, sheet_name=1)

    df_sales['ANNUAL PREMIUM'] = df_sales['ANNUAL PREMIUM'].apply(clean_currency)
    df_sales['TARGET PREMIUM'] = df_sales['TARGET PREMIUM'].apply(clean_currency)
    df_sales['DOANH SỐ THỰC'] = df_sales[['ANNUAL PREMIUM', 'TARGET PREMIUM']].min(axis=1)
    df_sales['MONTHS_DIFF'] = df_sales.apply(calculate_month_diff, axis=1)
    df_sales['ĐÁNH GIÁ TỐC ĐỘ'] = df_sales['MONTHS_DIFF'].apply(classify_speed)

    tab1, tab2, tab3 = st.tabs(["TỔNG QUAN", "NGUỒN SF", "NGUỒN CC"])

    with tab1:
        total_val = df_sales['DOANH SỐ THỰC'].sum()
        st.metric("TỔNG DOANH SỐ TOÀN TEAM", f"${total_val:,.2f}")
        
        c1, c2 = st.columns([3, 2])
        with c1:
            st.subheader("Doanh số theo Team")
            team_data = df_sales.groupby('TEAM')['DOANH SỐ THỰC'].sum().reset_index()
            st.bar_chart(team_data.set_index('TEAM'), color="#0f172a")
        with c2:
            st.subheader("Top Performance")
            owner_data = df_sales.groupby('OWNER')['DOANH SỐ THỰC'].sum().sort_values(ascending=False).reset_index()
            st.dataframe(owner_data, use_container_width=True, hide_index=True)

    with tab2:
        df_sf = df_sales[df_sales['SOURCE'] == 'SF']
        
        st.subheader("Chỉ số chuyển đổi")
        l_count = df_leads.groupby('OWNER')['LEAD ID'].count().reset_index(name='Nhận')
        c_count = df_sf.groupby('OWNER')['LEAD ID'].count().reset_index(name='Chốt')
        conv = pd.merge(l_count, c_count, on='OWNER', how='left').fillna(0)
        conv['Tỉ lệ %'] = (conv['Chốt'] / conv['Nhận'] * 100).round(2)
        st.dataframe(conv, use_container_width=True, hide_index=True)

        col_sf1, col_sf2 = st.columns(2)
        with col_sf1:
            st.subheader("Thống kê sản phẩm SF")
            st.dataframe(df_sf.groupby('PRODUCT')['LEAD ID'].count().reset_index(name='Số lượng'), use_container_width=True, hide_index=True)
        with col_sf2:
            st.subheader("Phân bổ tốc độ chốt")
            st.dataframe(df_sf['ĐÁNH GIÁ TỐC ĐỘ'].value_counts().reset_index(), use_container_width=True, hide_index=True)

        st.subheader("Chi tiết dữ liệu SF")
        st.dataframe(df_sf[['OWNER', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC', 'ĐÁNH GIÁ TỐC ĐỘ']], use_container_width=True, hide_index=True)

    with tab3:
        df_cc = df_sales[df_sales['SOURCE'] == 'CC']
        
        c_cc1, c_cc2 = st.columns([1, 2])
        with c_cc1:
            st.subheader("Thống kê sản phẩm CC")
            st.dataframe(df_cc.groupby('PRODUCT')['LEAD ID'].count().reset_index(name='Số lượng'), use_container_width=True, hide_index=True)
        with c_cc2:
            st.subheader("Chi tiết dữ liệu CC")
            st.dataframe(df_cc[['OWNER', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC']], use_container_width=True, hide_index=True)

    st.sidebar.markdown("---")
    if st.sidebar.button("Export Final Report"):
        df_sales.to_excel("Sales_Report.xlsx", index=False)
        st.sidebar.success("File Ready")
else:
    st.info("Please upload data file to start.")
