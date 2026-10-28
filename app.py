import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import base64
import os

st.set_page_config(page_title="Executive Dashboard", layout="wide")
C1, C2 = "#2E7D32", "#C8E6C9"

st.markdown("""
<style>
.block-container { padding-top: 40px !important; }
.kpi-box { background-color: #C8E6C9; padding: 20px; border-radius: 10px; border-left: 5px solid #2E7D32; text-align: center; margin-bottom: 10px; }
div[data-baseweb='tag'], span[data-baseweb='tag'], [class*='stMultiSelect'] span { background-color: #C8E6C9 !important; color: #1b5e20 !important; border: none !important; border-radius: 4px !important; padding: 4px 8px !important; }
div[data-baseweb='tag'] button svg, [class*='stMultiSelect'] span button svg { fill: #1b5e20 !important; }
div[data-baseweb='select'] > div { border: 1px solid #2E7D32 !important; }
div[role='button'][tabindex='0'] svg { fill: #2E7D32 !important; }
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_data():
    try:
        xl = pd.ExcelFile("sales_data.xlsx")
        sn, sd, br, pr, ts, tq, ov, ag_op = xl.sheet_names, {}, {}, {}, 0.0, 0.0, 0.0, pd.DataFrame()
        for s in sn:
            df_raw = pd.read_excel(xl, sheet_name=s, skiprows=1)
            sd[s] = df_raw
            
            df = pd.read_excel(xl, sheet_name=s, skiprows=1)
            if df.empty: continue
            df_c = df.copy()
            df_c.columns = [str(c).strip().lower() for c in df_c.columns]
            sc = [c for c in df_c.columns if any(k in c for k in ['sales', 'value', 'amount'])]
            qc = [c for c in df_c.columns if any(k in c for k in ['qty', 'quantity', 'count'])]
            
            if s.strip().lower() == "salessummary" and sc:
                v = pd.to_numeric(df_c.iloc[-1][sc], errors='coerce').dropna()
                if not v.empty: ts = float(v.iloc[0])
                bc = [c for c in df_c.columns if any(k in c for k in ['branch', 'zone', 'location', 'state'])]
                if bc:
                    for _, row in df_c.iterrows():
                        n = str(row[bc].iloc[0] if hasattr(row[bc], 'iloc') else row[bc]).upper().strip()
                        v_br = pd.to_numeric(row[sc], errors='coerce').dropna()
                        if not v_br.empty and n not in ['NAN', '', 'GRAND TOTAL', 'TOTAL']: br[n] = br.get(n, 0.0) + float(v_br.iloc[0])
            if s.strip().lower() == "dailysalesreport" and qc:
                v = pd.to_numeric(df_c.iloc[-1][qc], errors='coerce').dropna()
                if not v.empty: tq = float(v.iloc[0])
            if s.strip().replace(" ", "").lower() == "agingsummary_operatingbranch":
                try:
                    r_df = pd.read_excel(xl, sheet_name=s, skiprows=1)
                    r_df.columns = [str(c).strip() for c in r_df.columns]
                    r_map = {c: 'Branch' if 'branch' in str(c).lower() else 'Not Over Due' if 'not over' in str(c).lower() else 'Amount' if any(k in str(c).lower() for k in ['amount', 'value', 'balance', 'overdue']) else c for c in r_df.columns}
                    ag_op = r_df.rename(columns=r_map)
                    if 'Branch' in ag_op.columns:
                        ag_op = ag_op[ag_op['Branch'].notna() & (~ag_op['Branch'].astype(str).str.contains('Total|GRAND', case=False, na=False))]
                        for c in ['Amount', 'Not Over Due']:
                            if c in ag_op.columns: ag_op[c] = pd.to_numeric(ag_op[c], errors='coerce').fillna(0)
                    
                    rdf = pd.read_excel(xl, sheet_name=s, header=None)
                    if rdf.shape[0] >= 30 and rdf.shape[1] >= 10:
                        v_cell = rdf.iloc[29, 9]
                        v_num = pd.to_numeric(v_cell, errors='coerce')
                        if pd.notna(v_num): ov = float(v_num)
                except: ov = 0.0
            if s.strip().lower() == "salespdtsummary" and sc:
                pc = [c for c in df_c.columns if any(k in c for k in ['product', 'item', 'category', 'fly'])]
                if pc:
                    p_col, s_col = pc[0], sc[0]
                    for _, row in df_c.iterrows():
                        n = str(row[p_col]).strip()
                        v_p = pd.to_numeric(row[s_col], errors='coerce')
                        if pd.notna(v_p) and n not in ['nan', '', 'GRAND TOTAL', 'TOTAL']: pr[n] = pr.get(n, 0.0) + float(v_p)
                        
        b_df = pd.DataFrame(list(br.items()), columns=['Branch', 'Net Sales']).sort_values(by='Net Sales', ascending=False)
        p_df = pd.DataFrame(list(pr.items()), columns=['Product Category', 'Total Value']).sort_values(by='Total Value', ascending=False)
        return sd, sn, ts, tq, ov, b_df, p_df, ag_op
    except Exception as e:
        st.error(f"Excel Error: {e}")
        return None, [], 0.0, 0.0, 0.0, pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

all_sheets, sheet_names, total_sales, total_qty, total_overdue, branch_df, product_df, aging_operating_df = load_data()

if all_sheets:
    mode = st.sidebar.radio("View:", ["Summary", "Sheet View"])
    if mode == "Summary":
        if os.path.exists("FT_Logo.png"):
            with open("FT_Logo.png", "rb") as f: encoded_string = base64.b64encode(f.read()).decode()
            st.markdown(f'<div style="display: flex; justify-content: center; align-items: center; width: 100%; margin-top: 0px; margin-bottom: 0px;"><img src="data:image/png;base64,{encoded_string}" style="max-width: 280px; height: auto;"></div>', unsafe_allow_html=True)
        st.markdown(f"<h1 style='text-align: center; color: {C1}; margin-top: -10px; margin-bottom: 15px;'>Monthly Performance Dashboard</h1>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        col1.markdown(f"<div class='kpi-box'><h3>Total Sales</h3><h2>₹ {total_sales:,.2f}</h2></div>", unsafe_allow_html=True)
        col2.markdown(f"<div class='kpi-box'><h3>Total Qty</h3><h2>{total_qty:,.0f} Pcs</h2></div>", unsafe_allow_html=True)
        col3.markdown(f"<div class='kpi-box'><h3>Overdue Amount</h3><h2>₹ {total_overdue:,.2f}</h2></div>", unsafe_allow_html=True)
        g1, g2 = st.columns(2)
        if not branch_df.empty: g1.plotly_chart(px.bar(branch_df, x='Branch', y='Net Sales', text_auto='.2s', color_discrete_sequence=[C1], template="plotly_white", title="<b>Net Sales by Branch</b>"), use_container_width=True)
        if not product_df.empty: g2.plotly_chart(px.pie(product_df.head(10), names='Product Category', values='Total Value', hole=0.3, color_discrete_sequence=px.colors.sequential.Greens_r, title="<b>Sales Share by Product Category (Top 10)</b>"), use_container_width=True)
        st.markdown("<hr>", unsafe_allow_html=True)
        if not aging_operating_df.empty and 'Branch' in aging_operating_df.columns and 'Amount' in aging_operating_df.columns:
            plot_df = aging_operating_df.sort_values(by='Amount', ascending=False); fig_combo = go.Figure()
            fig_combo.add_trace(go.Bar(x=plot_df['Branch'], y=plot_df['Amount'], name='Overdue Amount', marker_color='#E53935', text=plot_df['Amount'].map('{:,.0f}'.format), textposition='auto'))
            if 'Not Over Due' in plot_df.columns: fig_combo.add_trace(go.Scatter(x=plot_df['Branch'], y=plot_df['Not Over Due'], name='Not Over Due', mode='lines+markers', line=dict(color='#1E88E5', width=3), marker=dict(size=8)))
            fig_combo.update_layout(title="<b>Customer Overdue Summary</b>", template="plotly_white", barmode='group', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(fig_combo, use_container_width=True)
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown(f"<h2 style='text-align: center; color: {C1};'>Stock Aging Summary</h2>", unsafe_allow_html=True)
        if "StockAgingSummary" in all_sheets:
            aging_df = all_sheets["StockAgingSummary"].copy(); aging_df.columns = [str(c).strip() for c in aging_df.columns]
            rename_map = {c: 'Branch' if 'branch' in str(c).lower() else 'AgeDays' if 'age' in str(c).lower() else 'PDT Fly' if any(k in str(c).lower() for k in ['pdt', 'fly']) else c for c in aging_df.columns}
            aging_df = aging_df.rename(columns=rename_map)
            val_cols = [c for c in aging_df.columns if any(k in str(c).lower() for k in ['qty', 'val', 'amount', 'balance', 'pcs']) or aging_df[c].dtype in ['int64', 'float64']]
            val_cols = [c for c in val_cols if c not in ['Branch', 'AgeDays', 'PDT Fly']]
            if 'Branch' in aging_df.columns and val_cols:
                clean_aging = aging_df[aging_df['Branch'].notna() & (~aging_df['Branch'].astype(str).str.contains('Total|GRAND', case=False, na=False))].copy()
                for vc in val_cols: clean_aging[vc] = pd.to_numeric(clean_aging[vc], errors='coerce').fillna(0)
                summary_aging = clean_aging.groupby('Branch')[val_cols].sum().reset_index().sort_values(by=val_cols, ascending=False)
                fig_age = px.bar(summary_aging, x='Branch', y=val_cols, title="<b>Stock Breakdown by Branch</b>", barmode='group', template="plotly_white")
                st.plotly_chart(fig_age, use_container_width=True)
            else: st.dataframe(aging_df.head(10), use_container_width=True)
        else: st.warning("Stock Aging Summary sheet could not be located.")
        
    elif mode == "Sheet View":
        st.markdown(f"<h2 style='color: {C1}; text-align: center;'>Workbook Raw Sheet Explorer</h2>", unsafe_allow_html=True)
        selected_sheet = st.selectbox("Choose a datasheet to examine:", sheet_names)
        if selected_sheet:
            current_df = all_sheets[selected_sheet].copy()
            st.markdown("#### 🔍 Filter Table Options")
            search_query = st.text_input("Global Search (Type anything to filter across all columns):", "")
            
            if search_query: current_df = current_df[current_df.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)]
            st.markdown("---")
            st.markdown(f"### Visualizing: {selected_sheet}")
            st.markdown("<style>tr:nth-child(even) {background-color: #F1F8E9 !important;} th {background-color: #2E7D32 !important; color: white !important;}</style>", unsafe_allow_html=True)
            st.dataframe(current_df.style.apply(lambda x: ["background-color: #F1F8E9" if idx % 2 == 0 else "background-color: #FFFFFF" for idx in range(len(x))], axis=0), use_container_width=True)



