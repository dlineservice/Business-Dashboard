import pandas as pd, plotly.express as px, plotly.graph_objects as go, streamlit as st, base64, os
st.set_page_config(page_title="Executive Dashboard", layout="wide")
C1, C2 = "#2E7D32", "#C8E6C9"
st.markdown(f"""<style>.block-container {{ padding-top: 40px !important; }} .kpi-box {{ background-color: {C2}; padding: 20px; border-radius: 10px; border-left: 5px solid {C1}; text-align: center; margin-bottom: 10px; }}</style>""", unsafe_allow_html=True)

@st.cache_data
def load_all_sheets():
    try:
        xl = pd.ExcelFile("sales_data.xlsx")
        return {s: pd.read_excel(xl, sheet_name=s) for s in xl.sheet_names}, xl.sheet_names
    except Exception as e:
        st.error(f"Excel Loading Error: {e}"); return {}, []

all_sheets, sheet_names = load_all_sheets()
if all_sheets:
    mode = st.sidebar.radio("View:", ["Summary", "Sheet View"])
    if mode == "Summary":
        ts, tq, ov, br, pr, aging_operating_df = 0.0, 0.0, 0.0, {}, {}, pd.DataFrame()
        for s, raw_df in all_sheets.items():
            if raw_df.empty or len(raw_df) <= 1: continue
            df = raw_df.iloc[1:].copy(); df.columns = [str(c).strip().lower() for c in raw_df.columns]
            sc = [c for c in df.columns if any(k in c for k in ['sales', 'value', 'amount'])]
            qc = [c for c in df.columns if any(k in c for k in ['qty', 'quantity', 'count'])]
            if s.strip().lower() == "salessummary" and sc:
                v = pd.to_numeric(df.iloc[-1][sc[0]], errors='coerce')
                if pd.notna(v): ts = float(v)
                bc = [c for c in df.columns if any(k in c for k in ['branch', 'zone', 'location', 'state'])]
                if bc:
                    for _, row in df.iterrows():
                        n = str(row[bc[0]]).upper().strip()
                        v_br = pd.to_numeric(row[sc[0]], errors='coerce')
                        if pd.notna(v_br) and n not in ['NAN', '', 'GRAND TOTAL', 'TOTAL']: br[n] = br.get(n, 0.0) + float(v_br)
            if s.strip().lower() == "dailysalesreport" and qc:
                v = pd.to_numeric(df.iloc[-1][qc[0]], errors='coerce')
                if pd.notna(v): tq = float(v)
            if s.strip().replace(" ", "").lower() == "agingsummary_operatingbranch":
                try:
                    r_df = df.copy(); r_df.columns = [str(c).strip() for c in raw_df.columns]
                    r_map = {c: 'Branch' if 'branch' in str(c).lower() else 'Not Over Due' if 'not over' in str(c).lower() else 'Amount' if any(k in str(c).lower() for k in ['amount', 'value', 'balance', 'overdue']) else c for c in r_df.columns}
                    aging_operating_df = r_df.rename(columns=r_map)
                    if 'Branch' in aging_operating_df.columns:
                        aging_operating_df = aging_operating_df[aging_operating_df['Branch'].notna() & (~aging_operating_df['Branch'].astype(str).str.contains('Total|GRAND', case=False, na=False))]
                        for c in ['Amount', 'Not Over Due']:
                            if c in aging_operating_df.columns: aging_operating_df[c] = pd.to_numeric(aging_operating_df[c], errors='coerce').fillna(0)
                    if raw_df.shape[0] >= 30 and raw_df.shape[1] >= 10:
                        v_cell = raw_df.iloc[29, 9]
                        v_num = pd.to_numeric(v_cell, errors='coerce')
                        if pd.notna(v_num): ov = float(v_num)
                except: ov = 0.0
            if s.strip().lower() == "salespdtsummary" and sc:
                pc = [c for c in df.columns if any(k in c for k in ['product', 'item', 'category', 'fly'])]
                if pc:
                    for _, row in df.iterrows():
                        n = str(row[pc[0]]).strip(); v_p = pd.to_numeric(row[sc[0]], errors='coerce')
                        if pd.notna(v_p) and n not in ['nan', '', 'GRAND TOTAL', 'TOTAL']: pr[n] = pr.get(n, 0.0) + float(v_p)
        branch_df = pd.DataFrame(list(br.items()), columns=['Branch', 'Net Sales']).sort_values(by='Net Sales', ascending=False)
        product_df = pd.DataFrame(list(pr.items()), columns=['Product Category', 'Total Value']).sort_values(by='Total Value', ascending=False)
        if os.path.exists("FT_Logo.png"):
            with open("FT_Logo.png", "rb") as f: encoded_string = base64.b64encode(f.read()).decode()
            st.markdown(f"""<div style="display: flex; justify-content: center;"><img src="data:image/png;base64,{encoded_string}" style="max-width: 280px;"></div>""", unsafe_allow_html=True)
        st.markdown(f"<h1 style='text-align: center; color: {C1}; margin-top: -10px; margin-bottom: 15px;'>Monthly Performance Dashboard</h1>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        col1.markdown(f"<div class='kpi-box'><h3>Total Sales</h3><h2>₹ {ts:,.2f}</h2></div>", unsafe_allow_html=True)
        col2.markdown(f"<div class='kpi-box'><h3>Total Qty</h3><h2>{tq:,.0f} Pcs</h2></div>", unsafe_allow_html=True)
        col3.markdown(f"<div class='kpi-box'><h3>Overdue Amount</h3><h2>₹ {ov:,.2f}</h2></div>", unsafe_allow_html=True)
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
        target_sheet = [s for s in all_sheets.keys() if "aging" in s.lower() and "stock" in s.lower()]
        sheet_key = target_sheet[0] if target_sheet else ("StockAgingSummary" if "StockAgingSummary" in all_sheets else None)
        if sheet_key:
            aging_df = all_sheets[sheet_key].copy(); aging_df.columns = [str(c).strip() for c in aging_df.columns]
            rename_map = {c: 'Branch' if 'branch' in str(c).lower() else 'AgeDays' if 'age' in str(c).lower() else 'PDT Fly' if any(k in str(c).lower() for k in ['pdt', 'fly']) else c for c in aging_df.columns}
            aging_df = aging_df.rename(columns=rename_map)
            val_cols = [c for c in aging_df.columns if any(k in str(c).lower() for k in ['qty', 'val', 'amount', 'balance', 'pcs']) or aging_df[c].dtype in ['int64', 'float64']]
            val_cols = [c for c in val_cols if c not in ['Branch', 'AgeDays', 'PDT Fly']]
            if 'Branch' in aging_df.columns and val_cols:
                clean_aging = aging_df[aging_df['Branch'].notna() & (~aging_df['Branch'].astype(str).str.contains('Total|GRAND', case=False, na=False))].copy()
                for vc in val_cols: clean_aging[vc] = pd.to_numeric(clean_aging[vc], errors='coerce').fillna(0)
                summary_aging = clean_aging.groupby('Branch')[val_cols[0]].sum().reset_index().sort_values(by=val_cols[0], ascending=False)
                fig_age = px.bar(summary_aging, x='Branch', y=val_cols[0], title=f"<b>Stock Breakdown by Branch</b>", template="plotly_white", color_discrete_sequence=['#43A047'])
                st.plotly_chart(fig_age, use_container_width=True)
            else: st.dataframe(aging_df.head(10), use_container_width=True)
        else: st.warning("Stock Aging Summary sheet could not be located in the current workbook.")
    elif mode == "Sheet View":
        st.markdown(f"<h2 style='color: {C1}; text-align: center;'>Workbook Raw Sheet Explorer</h2>", unsafe_allow_html=True)
        selected_sheet = st.selectbox("Choose a raw datasheet to examine:", sheet_names)
        if selected_sheet in all_sheets:
            st.markdown(f"### Visualizing: {selected_sheet}")
            st.dataframe(all_sheets[selected_sheet], use_container_width=True)
else: st.error("No data found or sales_data.xlsx missing.")
