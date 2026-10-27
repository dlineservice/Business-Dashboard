import pandas as pd, plotly.express as px, streamlit as st
st.set_page_config(page_title="Executive Dashboard", layout="wide")
C1, C2 = "#2E7D32", "#C8E6C9"
st.markdown(f"<style>.kpi-box {{ background-color: {C2}; padding: 20px; border-radius: 10px; border-left: 5px solid {C1}; text-align: center; margin-bottom: 10px; }} div[data-baseweb='tag'], span[data-baseweb='tag'], [class*='stMultiSelect'] span {{ background-color: {C2} !important; color: #1b5e20 !important; border: none !important; border-radius: 4px !important; padding: 4px 8px !important; }} div[data-baseweb='tag'] button svg, [class*='stMultiSelect'] span button svg {{ fill: #1b5e20 !important; }} div[data-baseweb='select'] > div {{ border: 1px solid {C1} !important; }} div[role='button'][tabindex='0'] svg, svg[class*='baseweb-icon'] {{ fill: {C1} !important; }}</style>", unsafe_allow_html=True)

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
            if df.empty: continue
            df_c = df.copy(); df_c.columns = [str(c).strip().lower() for c in df_c.columns]
            sc = [c for c in df_c.columns if any(k in c for k in ['sales', 'value', 'amount'])]
            qc = [c for c in df_c.columns if any(k in c for k in ['qty', 'quantity', 'count'])]
            
            if s.strip().lower() == "salessummary" and sc:
                r = df_c.iloc[-1]
                v_series = pd.to_numeric(r[sc], errors='coerce')
                v = v_series.iloc[0] if hasattr(v_series, 'iloc') else v_series
                if pd.notna(v): ts = float(v)
                bc = [c for c in df_c.columns if any(k in c for k in ['branch', 'zone', 'location', 'state'])]
                if bc:
                    b_col = bc[0]
                    s_col = sc[0]
                    for _, row in df_c.iterrows():
                        n = str(row[b_col]).upper().strip()
                        v_br = pd.to_numeric(row[s_col], errors='coerce')
                        if pd.notna(v_br) and n not in ['NAN', '', 'GRAND TOTAL', 'TOTAL']: br[n] = br.get(n, 0.0) + float(v_br)
            if s.strip().lower() == "dailysalesreport" and qc:
                r = df_c.iloc[-1]
                v_series = pd.to_numeric(r[qc], errors='coerce')
                v = v_series.iloc[0] if hasattr(v_series, 'iloc') else v_series
                if pd.notna(v): tq = float(v)
            if s.strip().replace(" ", "").lower() == "agingsummary_operatingbranch":
                try:
                    rdf = pd.read_excel(xl, sheet_name=s, header=None)
                    if rdf.shape[0] >= 30 and rdf.shape[1] >= 10:
                        v_raw = rdf.iloc[29, 9]
                        v_num = pd.to_numeric(v_raw, errors='coerce')
                        if pd.notna(v_num): ov = float(v_num)
                except: ov = 0.0
            if s.strip().lower() == "salespdtsummary":
                pc = [c for c in df_c.columns if any(k in c for k in ['product', 'item', 'category', 'fly'])]
                if pc and sc:
                    p_col = pc[0]
                    s_col = sc[0]
                    for _, row in df_c.iterrows():
                        n = str(row[p_col]).strip()
                        v_p = pd.to_numeric(row[s_col], errors='coerce')
                        if pd.notna(v_p) and n not in ['nan', '', 'GRAND TOTAL', 'TOTAL']: pr[n] = pr.get(n, 0.0) + float(v_p)
        b_df = pd.DataFrame(list(br.items()), columns=['Branch', 'Net Sales']).sort_values(by='Net Sales', ascending=False)
        p_df = pd.DataFrame(list(pr.items()), columns=['Product Category', 'Total Value']).sort_values(by='Total Value', ascending=False)
        return sd, sn, ts, tq, (0.0 if pd.isna(ov) else ov), b_df, p_df
    except Exception as e:
        st.error(f"Excel Error: {e}"); return None, [], 0.0, 0.0, 0.0, pd.DataFrame(), pd.DataFrame()

all_sheets, sheet_names, total_sales, total_qty, total_overdue, branch_df, product_df = load_and_process_dashboard_data()

if all_sheets:
    st.sidebar.title("Control Panel")
    mode = st.sidebar.radio("View:", ["Summary", "Sheet View"])
    if mode == "Summary":
        # --- লোগো সেকশন ---
        logo_col1, logo_col2, logo_col3 = st.columns([1, 2, 1])
        with logo_col2:
            try:
                st.image("FT_Logo.png", width=300)
            except:
                pass
                
        st.markdown(f"<h1 style='text-align: center; color: {C1};'>Fronttech Performance Dashboard</h1>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        col1.markdown(f"<div class='kpi-box'><h3>Total Sales</h3><h2>₹ {total_sales:,.2f}</h2></div>", unsafe_allow_html=True)
        col2.markdown(f"<div class='kpi-box'><h3>Total Qty</h3><h2>{total_qty:,.0f} Pcs</h2></div>", unsafe_allow_html=True)
        col3.markdown(f"<div class='kpi-box'><h3>Overdue Amount</h3><h2>₹ {total_overdue:,.2f}</h2></div>", unsafe_allow_html=True)
        
        g1, g2 = st.columns(2)
        if not branch_df.empty: 
            fig_bar = px.bar(
                branch_df, x='Branch', y='Net Sales', text_auto='.2s', 
                color_discrete_sequence=[C1], template="plotly_white",
                title="<b>Net Sales by Branch</b>"
            )
            g1.plotly_chart(fig_bar, use_container_width=True)
            
        if not product_df.empty: 
            fig_pie = px.pie(
                product_df.head(10), names='Product Category', values='Total Value', hole=0.3, 
                color_discrete_sequence=px.colors.sequential.Greens_r,
                title="<b>Sales Share by Product Category (Top 10)</b>"
            )
            g2.plotly_chart(fig_pie, use_container_width=True)
        
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown(f"<h2 style='text-align: center; color: {C1};'>Stock Aging Summary (Pivot Table)</h2>", unsafe_allow_html=True)
        
        if "StockAgingSummary" in all_sheets:
            aging_df = all_sheets["StockAgingSummary"].copy()
            aging_df.columns = [str(c).strip() for c in aging_df.columns]
            
            rename_map = {}
            for c in aging_df.columns:
                if 'branch' in str(c).lower(): rename_map[c] = 'Branch'
                elif 'age' in str(c).lower(): rename_map[c] = 'AgeDays'
                elif 'pdt' in str(c).lower() or 'fly' in str(c).lower(): rename_map[c] = 'PDT Fly'
            aging_df = aging_df.rename(columns=rename_map)
            
            if 'Branch' in aging_df.columns:
                aging_df = aging_df[aging_df['Branch'].notna() & (~aging_df['Branch'].astype(str).str.contains('Total|GRAND', case=False, na=False))]
                
                if 'AgeDays' in aging_df.columns:
                    aging_df['AgeDays'] = pd.to_numeric(aging_df['AgeDays'], errors='coerce').fillna(0)
                
                exclude_cols = ['Branch', 'AgeDays', 'PDT Fly']
                value_cols = [c for c in aging_df.columns if c not in exclude_cols]
                
                for c in value_cols:
                    aging_df[c] = pd.to_numeric(aging_df[c], errors='coerce').fillna(0)
                
                f1, f2 = st.columns(2)
                sel_branches = f1.multiselect("Filter Branch:", options=sorted(aging_df['Branch'].unique()), default=sorted(aging_df['Branch'].unique()))
                categories = sorted(aging_df['PDT Fly'].dropna().unique()) if 'PDT Fly' in aging_df.columns else []
                sel_categories = f2.multiselect("Filter PDT Fly:", options=categories, default=categories) if categories else []
                
                f_df = aging_df[aging_df['Branch'].isin(sel_branches)]
                if categories and sel_categories: 
                    f_df = f_df[f_df['PDT Fly'].isin(sel_categories)]
                
                if not f_df.empty and value_cols:
                    p_index = ['Branch']
                    if 'PDT Fly' in f_df.columns: p_index.append('PDT Fly')
                    
                    pivot_res = pd.pivot_table(f_df, values=value_cols, index=p_index, aggfunc='sum', margins=True, margins_name='Grand Total')
                    
                    st.dataframe(pivot_res, use_container_width=True)
                    st.download_button(label="Download Filtered CSV", data=f_df.to_csv(index=False).encode("utf-8"), file_name="StockAging_Filtered.csv", mime="text/csv")
                else: 
                    st.warning("No dynamic numeric value columns found or data is empty for selection.")
            else: 
                st.error("Column 'Branch' not found in StockAgingSummary sheet columns.")
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
