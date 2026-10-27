import pandas as pd, plotly.express as px, plotly.graph_objects as go, streamlit as st, base64, os
st.set_page_config(page_title="Executive Dashboard", layout="wide")
C1, C2 = "#2E7D32", "#C8E6C9"
st.markdown(f"""<style>.block-container {{ padding-top: 40px !important; }} .kpi-box {{ background-color: {C2}; padding: 20px; border-radius: 10px; border-left: 5px solid {C1}; text-align: center; margin-bottom: 10px; }} div[data-baseweb='tag'], span[data-baseweb='tag'], [class*='stMultiSelect'] span {{ background-color: {C2} !important; color: #1b5e20 !important; border: none !important; border-radius: 4px !important; padding: 4px 8px !important; }} div[data-baseweb='tag'] button svg, [class*='stMultiSelect'] span button svg {{ fill: #1b5e20 !important; }} div[data-baseweb='select'] > div {{ border: 1px solid {C1} !important; }} div[role='button'][tabindex='0'] svg {{ fill: {C1} !important; }}</style>""", unsafe_allow_html=True)
@st.cache_data
def load_data():
    try:
        xl = pd.ExcelFile("sales_data.xlsx")
        sn, sd, br, pr, ts, tq, ov, ag_op = xl.sheet_names, {}, {}, {}, 0.0, 0.0, 0.0, pd.DataFrame()
        for s in sn:
            df = pd.read_excel(xl, sheet_name=s, skiprows=1)
            sd[s] = df
            if df.empty: continue
            df_c = df.copy(); df_c.columns = [str(c).strip().lower() for c in df_c.columns]
            sc = [c for c in df_c.columns if any(k in c for k in ['sales', 'value', 'amount'])]
            qc = [c for c in df_c.columns if any(k in c for k in ['qty', 'quantity', 'count'])]
            if s.strip().lower() == "salessummary" and sc:
                v = pd.to_numeric(df_c.iloc[-1][sc], errors='coerce')
                if pd.notna(v).any(): ts = float(v.dropna().iloc)
                bc = [c for c in df_c.columns if any(k in c for k in ['branch', 'zone', 'location', 'state'])]
                if bc:
                    for _, row in df_c.iterrows():
                        n = str(row[bc].iloc[0] if hasattr(row[bc], 'iloc') else row[bc]).upper().strip(); v_br = pd.to_numeric(row[sc], errors='coerce').dropna()
                        if not v_br.empty and n not in ['NAN', '', 'GRAND TOTAL', 'TOTAL']: br[n] = br.get(n, 0.0) + float(v_br.iloc[0])
            if s.strip().lower() == "dailysalesreport" and qc:
                v = pd.to_numeric(df_c.iloc[-1][qc], errors='coerce')
                if pd.notna(v).any(): tq = float(v.dropna().iloc)
            if s.strip().replace(" ", "").lower() == "agingsummary_operatingbranch":
                try:
                    r_df = pd.read_excel(xl, sheet_name=s, skiprows=1); r_df.columns = [str(c).strip() for c in r_df.columns]
                    r_map = {c: 'Branch' if 'branch' in str(c).lower() else 'Not Over Due' if 'not over' in str(c).lower() else 'Amount' if any(k in str(c).lower() for k in ['amount', 'value', 'balance', 'overdue']) else c for c in r_df.columns}
                    ag_op = r_df.rename(columns=r_map)
                    if 'Branch' in ag_op.columns:
                        ag_op = ag_op[ag_op['Branch'].notna() & (~ag_op['Branch'].astype(str).str.contains('Total|GRAND', case=False, na=False))]
                        for c in ['Amount', 'Not Over Due']:
                            if c in ag_op.columns: ag_op[c] = pd.to_numeric(ag_op[c], errors='coerce').fillna(0)
                    rdf = pd.read_excel(xl, sheet_name=s, header=None)
                    if rdf.shape[0] >= 30 and rdf.shape[1] >= 10:
                        v_num = pd.to_numeric(rdf.iloc[29, 9], errors='coerce')
                        if pd.notna(v_num): ov = float(v_num)
                except: ov = 0.0
            if s.strip().lower() == "salespdtsummary" and sc:
                pc = [c for c in df_c.columns if any(k in c for k in ['product', 'item', 'category', 'fly'])]
                if pc:
                    for _, row in df_c.iterrows():
                        n = str(row[pc].iloc[0] if hasattr(row[pc], 'iloc') else row[pc]).strip(); v_p = pd.to_numeric(row[sc], errors='coerce').dropna()
                        if not v_p.empty and n not in ['nan', '', 'GRAND TOTAL', 'TOTAL']: pr[n] = pr.get(n, 0.0) + float(v_p.iloc[0])
        return sd, sn, ts, tq, ov, pd.DataFrame(list(br.items()), columns=['Branch', 'Net Sales']), pd.DataFrame(list(pr.items()), columns=['Product Category', 'Total Value']), ag_op
    except Exception as e:
        st.error(f"Excel Error: {e}"); return None, [], 0.0, 0.0, 0.0, pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
all_sheets, sheet_names, total_sales, total_qty, total_overdue, branch_df, product_df, aging_operating_df = load_data()
if all_sheets:
    mode = st.sidebar.radio("View:", ["Summary", "Sheet View"])
    if mode == "Summary":
        if os.path.exists("FT_Logo.png"):
            with open("FT_Logo.png", "rb") as f: encoded_string = base64.b64encode(f.read()).decode()
            st.markdown(f"""<div style="display: flex; justify-content: center; align-items: center; width: 100%; margin-top: 0px; margin-bottom: 0px;"><img src="data:image/png;base64,{encoded_string}" style="max-width: 280px; height: auto;"></div>""", unsafe_allow_html=True)
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
        st.markdown(f"<h2 style='text-align: center; color: {C1};'>Stock Aging Summary (Pivot Table)</h2>", unsafe_allow_html=True)
        if "StockAgingSummary" in all_sheets:
            aging_df = all_sheets["StockAgingSummary"].copy(); aging_df.columns = [str(c).strip() for c in aging_df.columns]
            rename_map = {c: 'Branch' if 'branch' in str(c).lower() else 'AgeDays' if 'age' in str(c).lower() else 'PDT Fly' if any(k in str(c).lower() for k in ['pdt', 'fly']) else c for c in aging_df.columns}
            aging_df = aging_df.rename(columns=rename_map)
            if 'Branch' in aging_df.columns:
                aging_df = aging_df[aging_df['Branch'].notna() & (~aging_df['Branch'].astype(str).str.contains('Total|GRAND', case=False, na=False))]
                if 'AgeDays' in aging_df.columns: aging_df['AgeDays'] = pd.to_numeric(aging_df['AgeDays'], errors='coerce').fillna(0)
                v_cols = [c for c in aging_df.columns if c not in ['Branch', 'AgeDays', 'PDT Fly']]
                for c in v_cols: aging_df[c] = pd.to_numeric(aging_df[c], errors='coerce').fillna(0)
                f1, f2 = st.columns(2)
                sel_branches = f1.multiselect("Filter Branch:", options=sorted(aging_df['Branch'].unique()), default=sorted(aging_df['Branch'].unique()))
                categories = sorted(aging_df['PDT Fly'].dropna().unique()) if 'PDT Fly' in aging_df.columns else []
                sel_categories = f2.multiselect("Filter PDT Fly:", options=categories, default=categories) if categories else []
                f_df = aging_df[aging_df['Branch'].isin(sel_branches)]
                if categories and sel_categories: f_df = f_df[f_df['PDT Fly'].isin(sel_categories)]
                if not f_df.empty and v_cols: st.dataframe(pd.pivot_table(f_df, values=v_cols, index=['Branch', 'PDT Fly'] if 'PDT Fly' in f_df.columns else ['Branch'], aggfunc='sum', margins=True, margins_name='Grand Total'), use_container_width=True)
    else:
        st.title("Sheet View")
        sh = st.selectbox("Select Sheet:", options=sheet_names)
        df_sh = pd.read_excel("sales_data.xlsx", sheet_name=sh, skiprows=1)
        if len(df_sh.columns) > 0:
            fc = st.selectbox("Filter Column:", options=df_sh.columns)
            f_df = df_sh[df_sh[fc].isin(st.multiselect("Select Values:", options=df_sh[fc].unique(), default=df_sh[fc].unique()))]
            st.dataframe(f_df, use_container_width=True)
            st.download_button(label="Download CSV", data=f_df.to_csv(index=False).encode("utf-8"), file_name=f"{sh}.csv", mime="text/csv")
