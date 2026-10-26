import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Executive Dashboard", layout="wide")
C1, C2 = "#2E7D32", "#C8E6C9"

st.markdown(f"""
    <style>
    .kpi-box {{
        background-color: {C2}; padding: 20px; border-radius: 10px; border-left: 5px solid {C1}; text-align: center; margin-bottom: 10px;
    }}
    div[data-baseweb="tag"], span[data-baseweb="tag"], [class*="stMultiSelect"] span, div[class*="stMultiSelect"] div span, [data-baseweb="tag"] *, span[data-baseweb="tag"] * {{
        border: none !important;
        border-style: none !important;
        outline: none !important;
        box-shadow: none !important;
    }}
    div[data-baseweb="tag"], span[data-baseweb="tag"], [class*="stMultiSelect"] span, div[class*="stMultiSelect"] div span {{
        background-color: {C2} !important;
        color: #1b5e20 !important;
        border-radius: 4px !important;
        padding: 4px 8px !important;
    }}
    div[data-baseweb="tag"] button svg, span[data-baseweb="tag"] button svg, [class*="stMultiSelect"] span button svg {{
        fill: #1b5e20 !important;
    }}
    div[data-baseweb="select"] > div {{
        border: 1px solid {C1} !important;
    }}
    div[role="button"][tabindex="0"] svg, svg[class*="baseweb-icon"] {{
        fill: {C1} !important;
    }}
    </style>
""", unsafe_allow_html=True)

@st.cache_data
def load_and_process_dashboard_data():
    try:
        xl = pd.ExcelFile("sales_data.xlsx")
        sn = xl.sheet_names
        sd, br, pr = {}, {}, {}
        ts, tq, ov = 0.0, 0.0, 0.0
        
        for s in sn:
            df = pd.read_excel(xl, sheet_name=s, skiprows=1)
            sd[s] = df
            if df.empty: 
                continue
            df_c = df.copy()
            df_c.columns = [str(c).strip().lower() for c in df_c.columns]
            
            sc = [c for c in df_c.columns if 'sales' in c or 'value' in c or 'amount' in c]
            qc = [c for c in df_c.columns if 'qty' in c or 'quantity' in c or 'count' in c]
            
            if s.strip().lower() == "salessummary":
                r = df_c.iloc[-1]
                if sc:
                    v_series = pd.to_numeric(r[sc], errors='coerce')
                    v = v_series.iloc[0] if hasattr(v_series, 'iloc') else v_series
                    if pd.notna(v): ts = float(v)
                if ts == 0.0 and len(df_c) > 1:
                    r = df_c.iloc[-2]
                    if sc:
                        v_series = pd.to_numeric(r[sc], errors='coerce')
                        v = v_series.iloc[0] if hasattr(v_series, 'iloc') else v_series
                        if pd.notna(v): ts = float(v)
                        
                bc = [c for c in df_c.columns if 'branch' in c or 'zone' in c or 'location' in c or 'state' in c]
                if bc and sc:
                    b_col_name = bc[0]
                    s_col_name = sc[0]
                    for idx, row in df_c.iterrows():
                        n = str(row[b_col_name]).upper().strip()
                        v = pd.to_numeric(row[s_col_name], errors='coerce')
                        if pd.notna(v) and n not in ['NAN', '', 'GRAND TOTAL', 'TOTAL']:
                            br[n] = br.get(n, 0.0) + float(v)
                            
            if s.strip().lower() == "dailysalesreport":
                r = df_c.iloc[-1]
                if qc:
                    v_series = pd.to_numeric(r[qc], errors='coerce')
                    v = v_series.iloc[0] if hasattr(v_series, 'iloc') else v_series
                    if pd.notna(v): tq = float(v)
                if tq == 0.0 and len(df_c) > 1:
                    r = df_c.iloc[-2]
                    if qc:
                        v_series = pd.to_numeric(r[qc], errors='coerce')
                        v = v_series.iloc[0] if hasattr(v_series, 'iloc') else v_series
                        if pd.notna(v): tq = float(v)
                        
            if s.strip().replace(" ", "").lower() == "agingsummary_operatingbranch":
                try:
                    rdf = pd.read_excel(xl, sheet_name=s, header=None)
                    if rdf.shape[0] >= 30 and rdf.shape[1] >= 10:
                        v_raw = rdf.iloc[29, 9]
                        v = pd.to_numeric(v_raw, errors='coerce')
                        if pd.notna(v): ov = float(v)
                except: 
                    ov = 0.0
                    
            if s.strip().lower() == "salespdtsummary":
                pc = [c for c in df_c.columns if 'product' in c or 'item' in c or 'category' in c or 'fly' in c]
                if pc and sc:
                    p_col_name = pc[0]
                    s_col_name = sc[0]
                    for idx, row in df_c.iterrows():
                        n = str(row[p_col_name]).strip()
                        v = pd.to_numeric(row[s_col_name], errors='coerce')
                        if pd.notna(v) and n not in ['nan', '', 'GRAND TOTAL', 'TOTAL']:
                            pr[n] = pr.get(n, 0.0) + float(v)
                            
        b_df = pd.DataFrame(list(br.items()), columns=['Branch', 'Net Sales']).sort_values(by='Net Sales', ascending=False)
        p_df = pd.DataFrame(list(pr.items()), columns=['Product Category', 'Total Value']).sort_values(by='Total Value', ascending=False)
        return sd, sn, ts, tq, (0.0 if pd.isna(ov) else ov), b_df, p_df
    except Exception as e:
        st.error(f"Excel Error: {e}")
        return None, [], 0.0, 0.0, 0.0, pd.DataFrame(), pd.DataFrame()

all_sheets, sheet_names, total_sales, total_qty, total_overdue, branch_df, product_df = load_and_process_dashboard_data()

if all_sheets:
    st.sidebar.title("Control Panel")
    mode = st.sidebar.radio("View:", ["Summary", "Sheet View"])
    if mode == "Summary":
        st.markdown(f"<h1 style='text-align: center; color: {C1};'>Performance Dashboard</h1>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        col1.markdown(f"<div class='kpi-box'><h3>Total Sales</h3><h2>₹ {total_sales:,.2f}</h2></div>", unsafe_allow_html=True)
        col2.markdown(f"<div class='kpi-box'><h3>Total Qty</h3><h2>{total_qty:,.0f} Pcs</h2></div>", unsafe_allow_html=True)
        col3.markdown(f"<div class='kpi-box'><h3>Overdue Amount</h3><h2>₹ {total_overdue:,.2f}</h2></div>", unsafe_allow_html=True)
        g1, g2 = st.columns(2)
        if not branch_df.empty: 
            g1.plotly_chart(px.bar(branch_df, x='Branch', y='Net Sales', text_auto='.2s', color_discrete_sequence=[C1], template="plotly_white"), use_container_width=True)
        if not product_df.empty: 
            g2.plotly_chart(px.pie(product_df.head(10), names='Product Category', values='Total Value', hole=0.3, color_discrete_sequence=px.colors.sequential.Greens_r), use_container_width=True)
            
        # --- StockAgingSummary Pivot Section ---
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown(f"<h2 style='text-align: center; color: {C1};'>Stock Aging Summary (Pivot Table)</h2>", unsafe_allow_html=True)
        
        if "StockAgingSummary" in all_sheets:
            raw_aging = all_sheets["StockAgingSummary"]
            header_row = 0
            for idx, row in raw_aging.iterrows():
                if 'branch' in [str(x).strip().lower() for x in row.values]:
                    header_row = idx + 1
                    break
                    
            aging_df = pd.read_excel("sales_data.xlsx", sheet_name="StockAgingSummary", skiprows=header_row)
            aging_df.columns = [str(c).strip() for c in aging_df.columns]
            
            if 'Branch' in aging_df.columns:
                aging_df = aging_df[aging_df['Branch'].notna() & (~aging_df['Branch'].str.contains('Total|GRAND', case=False, na=False))]
                
                aging_df['Qty'] = pd.to_numeric(aging_df['Qty'], errors='coerce').fillna(0)
                aging_df['Amount'] = pd.to_numeric(aging_df['Amount'], errors='coerce').fillna(0)
                aging_df['AgeDays'] = pd.to_numeric(aging_df['AgeDays'], errors='coerce').fillna(0)
                
                f1, f2, f3 = st.columns(3)
                branches = sorted(aging_df['Branch'].unique())
                sel_branches = f1.multiselect("Filter Branch:", options=branches, default=branches)
                
                categories = sorted(aging_df['PDT Fly'].dropna().unique()) if 'PDT Fly' in aging_df.columns else []
                sel_categories = f2.multiselect("Filter PDT Fly:", options=categories, default=categories) if categories else []
                
                max_days = int(aging_df['AgeDays'].max()) if not aging_df.empty else 365
                sel_days = f3.slider("Filter AgeDays Range:", min_value=0, max_value=max_value, value=(0, max_value))
                
                f_df = aging_df[
                    (aging_df['Branch'].isin(sel_branches)) & 
                    (aging_df['AgeDays'] >= sel_days) & 
                    (aging_df['AgeDays'] <= sel_days)
                ]
                if categories and sel_categories:
                    f_df = f_df[f_df['PDT Fly'].isin(sel_categories)]
                    
                if not f_df.empty:
                    p_index = ['Branch']
                    if 'PDT Fly' in f_df.columns:
                        p_index.append('PDT Fly')
                        
                    pivot_res = pd.pivot_table(
                        f_df,
                        values=['Qty', 'Amount'],
                        index=p_index,
                        aggfunc={'Qty': 'sum', 'Amount': 'sum'},
                        margins=True,
                        margins_name='Grand Total'
                    )
                    pivot_res = pivot_res[['Qty', 'Amount']]
                    
                    st.dataframe(pivot_res, use_container_width=True)
st.download_button(label="Download Filtered CSV", data=f_df.to_csv(index=False).encode("utf-8"), file_name="StockAging_Filtered.csv", mime="text/csv")
else:
st.warning("No data found matching the selected filters.")
else:
st.error("Sheet 'StockAgingSummary' not found in excel file.")
else:
st.title("Sheet View")
sh = st.selectbox("Select Sheet:", options=sheet_names)
df_sh = pd.read_excel("sales_data.xlsx", sheet_name=sh, skiprows=1)
st.subheader(f"Table: {sh}")
if len(df_sh.columns) > 0:
fc = st.selectbox("Filter Column:", options=df_sh.columns)
sv = st.multiselect("Select Values:", options=df_sh[fc].unique(), default=df_sh[fc].unique())
f_df = df_sh[df_sh[fc].isin(sv)]
st.dataframe(f_df, use_container_width=True)
st.download_button(label="Download CSV", data=f_df.to_csv(index=False).encode("utf-8"), file_name=f"{sh}.csv", mime="text/csv")
