import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Sales Manager", layout="wide")

# CSS chuyên nghiệp: Tăng khoảng cách, font Inter, màu tối sang trọng
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; background-color: #f8fafc; }
    .main { padding: 2rem 5rem; }
    h1 { font-weight: 800; color: #0f172a; letter-spacing: -0.05em; margin-bottom: 2rem; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px; }
    h3 { font-weight: 700; color: #1e293b; margin-top: 2rem; margin-bottom: 1rem; }
    [data-testid="stMetric"] { background-color: white; border: 1px solid #e2e8f0; padding: 25px !important; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
    [data-testid="stMetricValue"] { color: #0f172a; font-weight: 700 !important; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; background-color: #f1f5f9; padding: 8px; border-radius: 12px; }
    .stTabs [data-baseweb="tab"] { height: 45px; border-radius: 8px; border: none; padding: 0 30px; font-weight: 600; color: #64748b; }
    .stTabs [aria-selected="true"] { background-color: #0f172a !important; color: white !important; }
    .stDownloadButton > button { background-color: #0f172a !important; color: white !important; border-radius: 8px !important; width: 100%; padding: 12px; font-weight: 700; border: none !important; text-transform: uppercase; }
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
    # Đọc dữ liệu
    df_sales = pd.read_excel(uploaded_file, sheet_name=0)
    df_leads = pd.read_excel(uploaded_file, sheet_name=1)

    # 1. Logic Doanh số: min(Annual, Target)
    df_sales['ANNUAL PREMIUM'] = df_sales['ANNUAL PREMIUM'].apply(clean_currency)
    df_sales['TARGET PREMIUM'] = df_sales['TARGET PREMIUM'].apply(clean_currency)
    df_sales['DOANH SỐ THỰC'] = df_sales[['ANNUAL PREMIUM', 'TARGET PREMIUM']].min(axis=1)
    
    # 2. Logic Tốc độ
    df_sales['MONTHS_DIFF'] = df_sales.apply(calculate_diff, axis=1)
    def eval_speed(x):
        if pd.isna(x): return "Tự khai thác"
        return f"{int(x)} tháng - {'Chậm' if x > 6 else 'Nhanh'}"
    df_sales['ĐÁNH GIÁ'] = df_sales['MONTHS_DIFF'].apply(eval_speed)

    # Sidebar: Nút Export (Khắc phục lỗi logic thiếu nút)
    st.sidebar.markdown("---")
    export_buffer = io.BytesIO()
    with pd.ExcelWriter(export_buffer, engine='xlsxwriter') as writer:
        df_sales.to_excel(writer, index=False, sheet_name='Sales_Report')
    
    st.sidebar.download_button(
        label="Download Final Report",
        data=export_buffer.getvalue(),
        file_name=f"Manager_Report_{datetime.now().strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    tab1, tab2, tab3 = st.tabs(["TỔNG QUAN", "NGUỒN SF", "NGUỒN CC"])

    # --- TAB 1: TỔNG QUAN ---
    with tab1:
        st.metric("TỔNG DOANH SỐ TOÀN TEAM", f"${df_sales['DOANH SỐ THỰC'].sum():,.2f}")
        c1, c2 = st.columns([3, 2], gap="large")
        with c1:
            st.markdown("### Doanh số theo Team")
            t_data = df_sales.groupby('TEAM')['DOANH SỐ THỰC'].sum().reset_index()
            st.bar_chart(t_data.set_index('TEAM'), color="#0f172a")
        with c2:
            st.markdown("### Top Performance")
            o_data = df_sales.groupby('OWNER')['DOANH SỐ THỰC'].sum().sort_values(ascending=False).reset_index()
            st.dataframe(o_data, use_container_width=True, hide_index=True)

    # --- TAB 2: NGUỒN SF ---
    with tab2:
        df_sf = df_sales[df_sales['SOURCE'] == 'SF']
        
        # Logic chuyển đổi: Khớp Lead ID giữa 2 sheet
        st.markdown("### Tỉ lệ chuyển đổi Lead SF")
        df_leads['Month_Year'] = df_leads['DATE ADDED'].dt.strftime('%m/%Y')
        l_summary = df_leads.groupby(['OWNER', 'Month_Year'])['LEAD ID'].count().reset_index(name='Nhận')
        
        # Chỉ những lead ID trong Sheet 2 mà xuất hiện trong Sheet 1 (chốt)
        closed_ids = df_sf['LEAD ID'].unique()
        df_leads['Is_Closed'] = df_leads['LEAD ID'].isin(closed_ids)
        c_summary = df_leads[df_leads['Is_Closed'] == True].groupby(['OWNER', 'Month_Year'])['LEAD ID'].count().reset_index(name='Chốt')
        
        conv_final = pd.merge(l_summary, c_summary, on=['OWNER', 'Month_Year'], how='left').fillna(0)
        conv_final['%'] = (conv_final['Chốt'] / conv_final['Nhận'] * 100).round(2)
        st.dataframe(conv_final.sort_values(by=['Month_Year', '%'], ascending=False), use_container_width=True, hide_index=True)

        cs1, cs2 = st.columns(2, gap="medium")
        with cs1:
            st.markdown("### Sản phẩm SF")
            p_sf = df_sf.groupby('PRODUCT')['LEAD ID'].count().reset_index(name='Qty').sort_values(by='Qty', ascending=False)
            st.dataframe(p_sf, use_container_width=True, hide_index=True)
        with cs2:
            st.markdown("### Thống kê Tốc độ")
            st.dataframe(df_sf['ĐÁNH GIÁ'].value_counts().reset_index(name='Qty'), use_container_width=True, hide_index=True)

        st.markdown("### Chi tiết xử lý SF")
        # Sử dụng column_config để highlight "Chậm" mà không bị lỗi AttributeError
        st.dataframe(
            df_sf[['OWNER', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC', 'ĐÁNH GIÁ']],
            use_container_width=True,
            hide_index=True,
            column_config={
                "ĐÁNH GIÁ": st.column_config.TextColumn(
                    "ĐÁNH GIÁ TỐC ĐỘ",
                    help="Highlight nếu xử lý trên 6 tháng",
                    validate="Chậm"
                ),
                "DOANH SỐ THỰC": st.column_config.NumberColumn(format="$%.2f")
            }
        )

    # --- TAB 3: NGUỒN CC ---
    with tab3:
        df_cc = df_sales[df_sales['SOURCE'] == 'CC']
        cc1, cc2 = st.columns([1, 2], gap="medium")
        with cc1:
            st.markdown("### Sản phẩm CC")
            p_cc = df_cc.groupby('PRODUCT')['LEAD ID'].count().reset_index(name='Qty').sort_values(by='Qty', ascending=False)
            st.dataframe(p_cc, use_container_width=True, hide_index=True)
        with cc2:
            st.markdown("### Chi tiết chốt CC")
            st.dataframe(
                df_cc[['OWNER', 'LEAD ID', 'PRODUCT', 'DOANH SỐ THỰC']], 
                use_container_width=True, 
                hide_index=True,
                column_config={"DOANH SỐ THỰC": st.column_config.NumberColumn(format="$%.2f")}
            )
else:
    st.info("Vui lòng tải file để hệ thống bắt đầu phân tích.")
