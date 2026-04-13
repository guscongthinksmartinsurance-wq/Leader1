import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Sales Manager", layout="wide")

# CSS tinh chỉnh chiều sâu và sự chuyên nghiệp
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        background-color: #f1f5f9;
    }
    
    .main {
        padding: 1.5rem 3rem;
    }

    /* Định dạng Header */
    h1 {
        font-weight: 800;
        color: #0f172a;
        margin-bottom: 2rem;
        border-bottom: 2px solid #e2e8f0;
        padding-bottom: 1rem;
    }

    /* Styling cho Metric - Nhìn là thấy số ngay */
    [data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 25px !important;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.05);
    }

    [data-testid="stMetricValue"] {
        color: #0f172a;
        font-size: 2.2rem !important;
        font-weight: 700 !important;
    }

    /* Cân chỉnh lề và khoảng cách các bảng */
    div.block-container {
        padding-top: 2rem;
    }

    .stDataFrame {
        border-radius: 10px;
        border: 1px solid #e2e8f0;
    }

    /* Tabs hiện đại */
    .stTabs [data-baseweb="tab-list"] {
        gap: 12px;
        background-color: #e2e8f0;
        padding: 6px;
        border-radius: 10px;
        margin-bottom: 20px;
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 8px 30px;
        background-color: transparent;
        color: #475569;
        font-weight: 600;
        border: none;
    }

    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #0f172a !important;
        box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);
    }

    /* Nút Export nổi bật */
    .stDownloadButton > button {
        background-color: #1e293b !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 12px 24px !important;
        font-weight: 600 !important;
        width: 100%;
        margin-top: 10px;
    }

    .stDownloadButton > button:hover {
        background-color: #334155 !important;
        box-shadow: 0 10px 15px -3px rgb(0 0 0 / 0.1);
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

    # Xử lý logic doanh số
    df_sales['ANNUAL PREMIUM'] = df_sales['ANNUAL PREMIUM'].apply(clean_currency)
    df_sales['TARGET PREMIUM'] = df_sales['TARGET PREMIUM'].apply(clean_currency)
    df_sales['DOANH SỐ THỰC'] = df_sales[['ANNUAL PREMIUM', 'TARGET PREMIUM']].min(axis=1)
    df_sales['MONTHS_DIFF'] = df_sales.apply(calculate_month_diff, axis=1)
    df_sales['ĐÁNH GIÁ TỐC ĐỘ'] = df_sales['MONTHS_DIFF'].apply(classify_speed)

    # Sidebar: Nút Export
    st.sidebar.markdown("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_sales.to_excel(writer, index=False, sheet_name='Sales_Report')
    
    st.sidebar.download_button(
        label="DOWNLOAD FINAL REPORT",
        data=output.getvalue(),
        file_name=f"Henry_Sales_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    tab1, tab2, tab3 = st.tabs(["📊 TỔNG QUAN", "📁 NGUỒN SF", "💼 NGUỒN CC"])

    with tab1:
        total_sales = df_sales['DOANH SỐ THỰC'].sum()
        st.metric("TỔNG DOANH SỐ TOÀN TEAM", f"${total_sales:,.2f}")
        
        st.markdown("###")
        col_t1, col_t2 = st.columns([3, 2], gap="large")
        
        with col_t1:
            st.markdown("**Doanh số theo Team**")
            team_data = df_sales.groupby('TEAM')['DOANH SỐ THỰC'].sum().reset_index()
            st.bar_chart(team_data.set_index('TEAM'), color="#1e293b")
        
        with col_t2:
            st.markdown("**Top Performance**")
            owner_data = df_sales.groupby('OWNER')['DOANH SỐ THỰC'].sum().sort_values(ascending=False).reset_index()
            st.dataframe(owner_data, use_container_width=True, hide_index=True)

    with tab2:
        df_sf = df_sales[df_sales['SOURCE'] == 'SF']
        
        col_sf_top1, col_sf_top2 = st.columns([2, 1], gap="medium")
        with col_sf_top1:
            st.markdown("**Chỉ số chuyển đổi nhân viên**")
            l_count = df_leads.groupby('OWNER')['LEAD ID'].count().reset_index(name='Nhận')
            c_count = df_sf.groupby('OWNER')['LEAD ID'].count().reset_index(name='Chốt')
            conv = pd.merge(l_count, c_count, on='OWNER', how='left').fillna(0)
            conv['Tỉ lệ %'] = (conv['Chốt'] / conv['Nhận'] * 100).round(2)
            st.dataframe(conv.sort_values(by='Tỉ lệ %', ascending=False), use_container_width=True, hide_index=True)
        
        with col_sf_top2:
            st.markdown("**Sản phẩm SF chốt nhiều nhất**")
            prod_sf = df_sf.groupby('PRODUCT')['LEAD ID'].count().reset_index(name='Số lượng').sort_values(by='Số lượng', ascending=False)
            st.dataframe(prod_sf, use_container_width=True, hide_index=True)

        st.markdown("###")
        st.markdown("**Chi tiết dữ liệu SF**")
        st.dataframe(df_sf[['OWNER', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC', 'ĐÁNH GIÁ TỐC ĐỘ']], use_container_width=True, hide_index=True)

    with tab3:
        df_cc = df_sales[df_sales['SOURCE'] == 'CC']
        
        col_cc_1, col_cc_2 = st.columns([1, 2], gap="medium")
        with col_cc_1:
            st.markdown("**Sản phẩm CC chốt nhiều nhất**")
            prod_cc = df_cc.groupby('PRODUCT')['LEAD ID'].count().reset_index(name='Số lượng').sort_values(by='Số lượng', ascending=False)
            st.dataframe(prod_cc, use_container_width=True, hide_index=True)
            
        with col_cc_2:
            st.markdown("**Dữ liệu chi tiết CC**")
            st.dataframe(df_cc[['OWNER', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC']], use_container_width=True, hide_index=True)
else:
    st.info("Anh hãy upload file dữ liệu để hệ thống bắt đầu phân tích.")
