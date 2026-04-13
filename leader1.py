import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Sales Manager", layout="wide")

st.markdown("""
    <style>
    .main {
        background-color: #f5f7f9;
    }
    .stMetric {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    div[data-testid="stExpander"] {
        border: none;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        background-color: white;
        margin-bottom: 10px;
    }
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        background-color: #1e293b;
        color: white;
        border: none;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 20px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        background-color: transparent;
        border-radius: 4px 4px 0 0;
        gap: 1px;
        padding-top: 10px;
        padding-bottom: 10px;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff;
        color: #1e293b;
    }
    h1, h2, h3 {
        color: #1e293b;
        font-family: 'Inter', sans-serif;
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
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Doanh số theo Team")
            team_data = df_sales.groupby('TEAM')['DOANH SỐ THỰC'].sum().reset_index()
            st.bar_chart(team_data.set_index('TEAM'), color="#1e293b")
        with c2:
            st.subheader("Top Performance")
            owner_data = df_sales.groupby('OWNER')['DOANH SỐ THỰC'].sum().sort_values(ascending=False).head(5)
            st.dataframe(owner_data, use_container_width=True)

    with tab2:
        df_sf = df_sales[df_sales['SOURCE'] == 'SF']
        
        st.subheader("Chỉ số chuyển đổi")
        l_count = df_leads.groupby('OWNER')['LEAD ID'].count().reset_index(name='Nhận')
        c_count = df_sf.groupby('OWNER')['LEAD ID'].count().reset_index(name='Chốt')
        conv = pd.merge(l_count, c_count, on='OWNER', how='left').fillna(0)
        conv['Tỉ lệ %'] = (conv['Chốt'] / conv['Nhận'] * 100).round(2)
        st.dataframe(conv, use_container_width=True)

        st.subheader("Thống kê sản phẩm SF")
        st.write(df_sf.groupby('PRODUCT')['LEAD ID'].count().reset_index(name='Số lượng'))

        with st.expander("Chi tiết dữ liệu SF"):
            st.dataframe(df_sf[['OWNER', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC', 'ĐÁNH GIÁ TỐC ĐỘ']], use_container_width=True)

    with tab3:
        df_cc = df_sales[df_sales['SOURCE'] == 'CC']
        
        st.subheader("Thống kê sản phẩm CC")
        st.write(df_cc.groupby('PRODUCT')['LEAD ID'].count().reset_index(name='Số lượng'))
        
        st.subheader("Chi tiết dữ liệu CC")
        st.dataframe(df_cc[['OWNER', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC']], use_container_width=True)

    if st.sidebar.button("Export Final Report"):
        df_sales.to_excel("Sales_Report.xlsx", index=False)
        st.sidebar.success("File Ready")
else:
    st.info("Please upload data file to start.")