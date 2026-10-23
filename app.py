import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Executive Monthly Performance Dashboard", layout="wide")

PRIMARY_COLOR = "#2E7D32"
SECONDARY_COLOR = "#C8E6C9"

st.markdown(f"<style>.kpi-box {{background-color: {SECONDARY_COLOR}; padding: 20px; border-radius: 10px; border-left: 5px solid {PRIMARY_COLOR}; text-align: center; margin-bottom: 10px;}}</style>", unsafe_allow_html=True)

@st.cache_data
def load_and_process_dashboard_data():
    try:
        excel_file = pd.ExcelFile("sales_data.xlsx")
        sheet_names = excel_file.sheet_names
        all_sheets_dict = {}
        total_business_sales = 0.0
        total_business_qty = 0
        branch_summary_dict = {}
        product_summary_dict = {}
        
        for sheet in sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet, skiprows=1)
            all_sheets_dict[sheet] = df
            df_clean = df.copy()
            df_clean.columns = [str(col).strip().lower() for col in df_clean.columns]
            
            sales_cols = [c for c in df_clean.columns if 'sales' in c or 'value' in c or 'amount' in c]
            qty_cols = [c for c in df_clean.columns if 'qty' in c or 'quantity' in c or 'count' in c]
            
            # 🎯 ফিক্সড লজিক: শুধুমাত্র 'SalesSummary' শীটের শেষ লাইনের আসল Total রো থেকে ভ্যালু নেওয়া হচ্ছে
            if sheet.strip().lower() == "salessummary" and not df_clean.empty:
                # শেষ সারিটি খুঁজে বের করা
                last_row = df_clean.iloc[-1]
                
                # যদি একাধিক সেলস বা কোয়ান্টিটি কলাম থাকে তবে প্রথমটি নেওয়া হবে
                if sales_cols:
                    total_business_sales = pd.to_numeric(last_row[sales_cols[0]], errors='coerce')
                if qty_cols:
                    total_business_qty = pd.to_numeric(last_row[qty_cols[0]], errors='coerce')
                    
                # যদি কোনো কারণে শেষ রো-তে ভ্যালু না পাওয়া যায় (যেমন ফাঁকা লাইন), তবে তার আগেরটি চেক করবে
                if (pd.isna(total_business_sales) or total_business_sales == 0) and len(df_clean) > 1:
                    last_row = df_clean.iloc[-2]
                    if sales_cols:
                        total_business_sales = pd.to_numeric(last_row[sales_cols[0]], errors='coerce')
                    if qty_cols:
                        total_business_qty = pd.to_numeric(last_row[qty_cols[0]], errors='coerce')

            # ব্রাঞ্চ ভিত্তিক ডেটা প্রসেস
            branch_cols = [c for c in df_clean.columns if 'branch' in c or 'zone' in c or 'location' in c or 'state' in c]
            if branch_cols and sales_cols:
                b_col = branch_cols[0]
                s_col = sales_cols[0]
                for idx, row in df_clean.iterrows():
                    b_name = str(row[b_col]).upper().strip()
                    b_val = pd.to_numeric(row[s_col], errors='coerce')
                    if pd.notna(b_val) and b_name not in ['NAN', '', 'GRAND TOTAL', 'TOTAL']:
                        branch_summary_dict[b_name] = branch_summary_dict.get(b_name, 0) + b_val

            # প্রোডাক্ট ভিত্তিক ডেটা প্রসেস
            if sheet.strip().lower() == "salespdtsummary":
                pdt_cols = [c for c in df_clean.columns if 'product' in c or 'item' in c or 'category' in c or 'fly' in c]
                if pdt_cols and sales_cols:
                    p_col = pdt_cols[0]
                    s_col = sales_cols[0]
                    for idx, row in df_clean.iterrows():
                        p_name = str(row[p_col]).strip()
                        p_val = pd.to_numeric(row[s_col], errors='coerce')
                        if pd.notna(p_val) and p_name not in ['nan', '', 'GRAND TOTAL', 'TOTAL']:
                            product_summary_dict[p_name] = product_summary_dict.get(p_name, 0) + p_val

        # নাল ভ্যালু হ্যান্ডেল করা
        total_business_sales = 0.0 if pd.isna(total_business_sales) else total_business_sales
        total_business_qty = 0 if pd.isna(total_business_qty) else total_business_qty

        branch_df = pd.DataFrame(list(branch_summary_dict.items()), columns=['Branch', 'Net Sales']).sort_values(by='Net Sales', ascending=False)
        product_df = pd.DataFrame(list(product_summary_dict.items()), columns=['Product Category', 'Total Value']).sort_values(by='Total Value', ascending=False)
        return all_sheets_dict, sheet_names, total_business_sales, total_business_qty, branch_df, product_df
    except Exception as e:
        st.error(f"এক্সেল ফাইল প্রসেস করতে সমস্যা হচ্ছে: {e}")
        return None, [], 0, 0, pd.DataFrame(), pd.DataFrame()

all_sheets_dict, sheet_names, total_sales, total_qty, branch_df, product_df = load_and_process_dashboard_data()

if all_sheets_dict:
    st.sidebar.title("🟢 Executive Control Panel")
    st.sidebar.markdown("---")
    view_mode = st.sidebar.radio("পেজ ভিউ সিলেক্ট করুন:", ["🏠 মেইন ড্যাসবোর্ড (Summary)", "📂 ডিটেইলস শীট ভিউ (All 14 Sheets)"])
    st.sidebar.markdown("---")
    
    if view_mode == "🏠 মেইন ড্যাসবোর্ড (Summary)":
        st.markdown(f"<h1 style='text-align: center; color: {PRIMARY_COLOR};'>Northwind Traders Monthly Performance</h1>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; font-size: 16px;'>SalesSummary শীটের মূল Total রো এবং SalesPDTSummary শীটের ওপর ভিত্তি করে তৈরি সংক্ষিপ্ত রিপোর্ট</p>", unsafe_allow_html=True)
        st.markdown("---")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f"<div class='kpi-box'><h3 style='color: #1b5e20; margin:0;'>💰 Total Net Sales</h3><h2 style='color: #1b5e20; margin:10px 0;'>₹ {total_sales:,.2f}</h2><small>SalesSummary-এর শেষ লাইনের মূল Total ভ্যালু</small></div>", unsafe_allow_html=True)
        with col2:
            st.markdown(f"<div class='kpi-box'><h3 style='color: #1b5e20; margin:0;'>📦 Total Quantity Sold</h3><h2 style='color: #1b5e20; margin:10px 0;'>{total_qty:,.0f} Pcs</h2><small>SalesSummary-এর শেষ লাইনের মূল Total কোয়ান্টিটি</small></div>", unsafe_allow_html=True)
        with col3:
            st.markdown(f"<div class='kpi-box'><h3 style='color: #1b5e20; margin:0;'>📁 Total Active Sheets</h3><h2 style='color: #1b5e20; margin:10px 0;'>{len(sheet_names)} Sheets</h2><small>সফলভাবে সিঙ্ক হওয়া ডেটাসেট</small></div>", unsafe_allow_html=True)
            
        st.markdown("---")
        chart_col1, chart_col2 = st.columns(2)
        
        with chart_col1:
            st.markdown(f"<h3 style='color: {PRIMARY_COLOR};'>🏢 BRANCH ACTIVITY (Top Branches)</h3>", unsafe_allow_html=True)
            if not branch_df.empty:
                fig_branch = px.bar(branch_df.head(10), x='Branch', y='Net Sales', text_auto='.2s', color_discrete_sequence=[PRIMARY_COLOR], template="plotly_white")
                st.plotly_chart(fig_branch, use_container_width=True)
            else:
                st.info("ব্রাঞ্চ ভিত্তিক ডেটা প্রসেস করা সম্ভব হয়নি।")
                
        with chart_col2:
            st.markdown(f"<h3 style='color: {PRIMARY_COLOR};'>📦 PRODUCT ACTIVITY (Sales Share)</h3>", unsafe_allow_html=True)
            if not product_df.empty:
                fig_product = px.pie(product_df.head(10), names='Product Category', values='Total Value', hole=0.3, color_discrete_sequence=px.colors.sequential.Greens_r)
                st.plotly_chart(fig_product, use_container_width=True)
            else:
                st.info("SalesPDTSummary শীট থেকে প্রোডাক্ট ভিত্তিক ডেটা প্রসেস করা সম্ভব হয়নি।")

    elif view_mode == "📂 ডিটেইলস শীট ভিউ (All 14 Sheets)":
        st.title("📂 ইন্ডিভিজুয়াল শীট অ্যানালাইসিস ও ফিল্টারিং")
        selected_sheet = st.selectbox("📁 কোন শীটের ডিটেইলস ডেটা ও টেবিল দেখতে চান?", options=sheet_names)
        st.markdown("---")
        
        excel_file = pd.ExcelFile("sales_data.xlsx")
        df_sheet = pd.read_excel(excel_file, sheet_name=selected_sheet, skiprows=1)
        st.subheader(f"📋 টেবিল: {selected_sheet} (মোট ডেটা সারি: {len(df_sheet)})")
        
        if len(df_sheet.columns) > 0:
            filter_column = st.selectbox("🎯 ফিল্টার করার জন্য একটি কলাম বেছে নিন:", options=df_sheet.columns)
            unique_vals = df_sheet[filter_column].unique()
            selected_vals = st.multiselect(f"'{filter_column}' কলামের নির্দিষ্ট ভ্যালু ফিল্টার করুন:", options=unique_vals, default=unique_vals)
            
            filtered_df = df_sheet[df_sheet[filter_column].isin(selected_vals)]
            st.dataframe(filtered_df, use_container_width=True)
            
            csv = filtered_df.to_csv(index=False).encode("utf-8")
            st.download_button(label=f"📥 {selected_sheet} শীটের ফিল্টার করা ডেটা ডাউনলোড করুন", data=csv, file_name=f"{selected_sheet}_data.csv", mime="text/csv")
