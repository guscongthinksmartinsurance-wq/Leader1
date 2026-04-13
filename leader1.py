import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Sales Manager", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        background-color: #f8fafc;
    }

    .main { padding: 2rem 4rem; }

    /* Dashboard Header */
    h1 {
        font-weight: 800;
        color: #0f172a;
        font-size: 2.5rem;
        letter-spacing: -0.05em;
        margin-bottom: 2rem;
    }

    /* Thẻ chỉ số tổng */
    [data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 30px !important;
        border-radius: 16px;
        box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1);
    }

    /* Phân tách các bảng bằng khoảng trắng */
    .stTable, .stDataFrame {
        margin-top: 1.5rem;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
    }

    /* Tabs phong cách tối giản chuyên nghiệp */
    .stTabs [data-baseweb="tab-list"] {
        gap: 15px;
        background-color: transparent;
    }

    .stTabs [data-baseweb="tab"] {
        height: 50px;
        border-radius: 8px;
        background-color: #f1f5f9;
        color: #64748b;
        font-weight: 600;
        border: 1px solid #e2e8f0;
    }

    .stTabs [aria-selected="true"] {
        background-color: #1e293b !important;
        color: #ffffff !important;
    }

    /* Nút Download chuyên sâu */
    .stDownloadButton > button {
        background-color: #0f172a !important;
        color: white !important;
        border-radius: 8px !important;
        padding: 15px 30px !important;
        font-weight: 700 !important;
        width: 100%;
        border: none !important;
        text-transform: uppercase;
        letter-spacing: 1px;
    }

    /* Tạo khối cho từng phần */
    .section-box {
        background-color: white;
        padding: 25px;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
        margin-bottom: 2rem;
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
        if pd.isna(row['THÁNG NHẬN LEAD']) or pd.isna(row['NĂM NHẬN LEAD']): return None
        start_date = datetime(int(row['NĂM NHẬN LEAD']), int(row['THÁNG NHẬN LEAD']), 1)
        return (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month)
    except: return None

def classify_speed(m):
    if pd.isna(m): return "Tự khai thác"
    label = "Chậm" if m > 6 else "Nhanh"
    return f"{int(m)} tháng - {label}"

uploaded_file = st.sidebar.file_uploader("Upload Data File", type=["xlsx"])

if uploaded_file:
    df_sales = pd.read_excel(uploaded_file, sheet_name=0)
    df_leads = pd.read_excel(uploaded_file, sheet_name=1)

    df_sales['ANNUAL PREMIUM'] = df_sales['ANNUAL PREMIUM'].apply(clean_currency)
    df_sales['TARGET PREMIUM'] = df_sales['TARGET PREMIUM'].apply(clean_currency)
    df_sales['DOANH SỐ THỰC'] = df_sales[['ANNUAL PREMIUM', 'TARGET PREMIUM']].min(axis=1)
    df_sales['MONTHS_DIFF'] = df_sales.apply(calculate_month_diff, axis=1)
    df_sales['ĐÁNH GIÁ TỐC ĐỘ'] = df_sales['MONTHS_DIFF'].apply(classify_speed)

    # Nút Export nằm riêng biệt ở Sidebar
    st.sidebar.markdown("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_sales.to_excel(writer, index=False, sheet_name='Final_Report')
    
    st.sidebar.download_button(
        label="Download Final Report",
        data=output.getvalue(),
        file_name=f"Manager_Report_{datetime.now().strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    tab1, tab2, tab3 = st.tabs(["TỔNG QUAN", "NGUỒN SF", "NGUỒN CC"])

    with tab1:
        total = df_sales['DOANH SỐ THỰC'].sum()
        st.metric("TỔNG DOANH SỐ", f"${total:,.2f}")
        
        st.markdown("<br>", unsafe_allow_html=True)
        col1, col2 = st.columns([3, 2], gap="large")
        
        with col1:
            st.markdown("### Doanh số theo Team")
            t_data = df_sales.groupby('TEAM')['DOANH SỐ THỰC'].sum().reset_index()
            st.bar_chart(t_data.set_index('TEAM'), color="#0f172a")
        
        with col2:
            st.markdown("### Top Performance")
            o_data = df_sales.groupby('OWNER')['DOANH SỐ THỰC'].sum().sort_values(ascending=False).reset_index()
            st.dataframe(o_data, use_container_width=True, hide_index=True)

    with tab2:
        df_sf = df_sales[df_sales['SOURCE'] == 'SF']
        
        st.markdown("### Tỉ lệ chuyển đổi & Sản phẩm")
        c1, c2 = st.columns([2, 1], gap="medium")
        with c1:
            l_cnt = df_leads.groupby('OWNER')['LEAD ID'].count().reset_index(name='Nhận')
            c_cnt = df_sf.groupby('OWNER')['LEAD ID'].count().reset_index(name='Chốt')
            conv = pd.merge(l_cnt, c_cnt, on='OWNER', how='left').fillna(0)
            conv['Tỉ lệ %'] = (conv['Chốt'] / conv['Nhận'] * 100).round(2)
            st.dataframe(conv.sort_values(by='Tỉ lệ %', ascending=False), use_container_width=True, hide_index=True)
        
        with c2:
            p_sf = df_sf.groupby('PRODUCT')['LEAD ID'].count().reset_index(name='Qty')
            st.dataframe(p_sf.sort_values(by='Qty', ascending=False), use_container_width=True, hide_index=True)

        st.markdown("### Chi tiết xử lý SF")
        # Highlight logic cho bảng chi tiết
        def highlight_slow(s):
            return ['color: #ef4444; font-weight: bold' if 'Chậm' in str(v) else '' for v in s]
        
        st.dataframe(df_sf[['OWNER', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC', 'ĐÁNH GIÁ TỐC ĐỘ']]
                     .style.apply(highlight_slow, subset=['ĐÁNH GIÁ TỐC ĐỘ']), 
                     use_container_width=True, hide_index=True)

    with tab3:
        df_cc = df_sales[df_sales['SOURCE'] == 'CC']
        
        c3, c4 = st.columns([1, 2], gap="medium")
        with c3:
            st.markdown("### Sản phẩm CC")
            p_cc = df_cc.groupby('PRODUCT')['LEAD ID'].count().reset_index(name='Qty')
            st.dataframe(p_cc.sort_values(by='Qty', ascending=False), use_container_width=True, hide_index=True)
            
        with c4:
            st.markdown("### Chi tiết chốt CC")
            st.dataframe(df_cc[['OWNER', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC']], use_container_width=True, hide_index=True)

else:
    st.info("Hệ thống đang sẵn sàng. Anh vui lòng tải file lên.")
