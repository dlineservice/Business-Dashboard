import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Executive Dashboard", layout="wide")
C1, C2 = "#2E7D32", "#C8E6C9"

# Injecting clean global CSS to forcefully remove any internal borders and boxes from select items
st.markdown(f"""
    <style>
    .kpi-box {{
        background-color: {C2}; padding: 20px; border-radius: 10px; border-left: 5px solid {C1}; text-align: center; margin-bottom: 10px;
    }}
    /* Forcefully completely eliminate internal boxes, borders, and outlines from multiselect pills */
    div[data-baseweb="tag"], span[data-baseweb="tag"], [class*="stMultiSelect"] span, div[class*="stMultiSelect"] div span, [data-baseweb="tag"] *, span[data-baseweb="tag"] * {{
        border: none !important;
        border-style: none !important;
        outline: none !important;
        box-shadow: none !important;
    }}
    /* Apply a clean background to the outer pill container only */
    div[data-baseweb="tag"], span[data-baseweb="tag"], [class*="stMultiSelect"] span, div[class*="stMultiSelect"] div span {{
        background-color: {C2} !important;
        color: #1b5e20 !important;
        border-radius: 4px !important;
        padding: 4px 8px !important;
    }}
    div[data-baseweb="tag"] button svg, span[data-baseweb="tag"] button svg, [class*="stMultiSelect"] span button svg {{
        fill: #1b5e20 !important;
    }}
    /* Clean corporate green frame border for the main search input box only */
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
            if df.empty: continue
            df_c = df.copy()
            df_c.columns = [str(c).strip().lower() for c in df_c.columns]
            sc = [c for c in df_c.columns if 'sales' in c or 'value' in c or 'amount' in c]
            qc = [c for c in df_c.columns if 'qty' in c or 'quantity' in c or 'count' in c]
            if s.strip().lower() == "salessummary":
                r = df_c.iloc[-1]
                if sc:
                    v = pd.to_numeric(r[sc].values if hasattr(r[sc], "values") else r[sc], errors='coerce')
                    if pd.notna(v): ts = float(v)
                if ts == 0.0 and len(df_c) > 1:
                    r = df_c.iloc[-2]
                    if sc:
                        v = pd.to_numeric(r[sc].values if hasattr(r[sc], "values") else r[sc], errors='coerce')
                        if pd.notna(v): ts = float(v)
                bc = [c for c in df_c.columns if 'branch' in c or 'zone' in c or 'location' in c or 'state' in c]
                if bc and sc:
                    for i, row in df_c.iterrows():
                        n = str(row[bc].values if hasattr(row[bc], "values") else row[bc]).upper().strip()
                        v = pd.to_numeric(row[sc].values if hasattr(row[sc], "values") else row[sc], errors='coerce')
                        if pd.notna(v) and n not in ['NAN', '', 'GRAND TOTAL', 'TOTAL']: br[n] = br.get(n, 0.0) + float(v)
            if s.strip().lower() == "dailysalesreport":
                r = df_c.iloc[-1]
                if qc:
                    v = pd.to_numeric(r[qc].values if hasattr(r[qc], "values") else r[qc], errors='coerce')
                    if pd.notna(v): tq = float(v)
                if tq == 0.0 and len(df_c) > 1:
                    r = df_c.iloc[-2]
                    if qc:
                        v = pd.to_numeric(r[qc].values if hasattr(r[qc], "values") else r[qc], errors='coerce')
                        if pd.notna(v): tq = float(v)
            if s.strip().replace(" ", "").lower() == "agingsummary_operatingbranch":
                try:
                    rdf = pd.read_excel(xl, sheet_name=s, header=None)
                    if rdf.shape[0] >= 30 and rdf.shape[1] >= 10:
                        v_raw = rdf.iloc[29, 9]
                        v = pd.to_numeric(v_raw.values if hasattr(v_raw, "values") else v_raw, errors='coerce')
                        if pd.notna(v): ov = float(v)
                except: ov = 0.0
            if s.strip().lower() == "salespdtsummary":
                pc = [c for c in df_c.columns if 'product' in c or 'item' in c or 'category' in c or 'fly' in c]
                if pc and sc:
                    for i, row in df_c.iterrows():
                        n = str(row[pc].values if hasattr(row[pc], "values") else row[pc]).strip()
                        v = pd.to_numeric(row[sc].values if hasattr(row[sc], "values") else row[sc], errors='coerce')
                        if pd.notna(v) and n not in ['nan', '', 'GRAND TOTAL', 'TOTAL']: pr[n] = pr.get(n, 0.0) + float(v)
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
        if not branch_df.empty: g1.plotly_chart(px.bar(branch_df, x='Branch', y='Net Sales', text_auto='.2s', color_discrete_sequence=[C1], template="plotly_white"), use_container_width=True)
        if not product_df.empty: g2.plotly_chart(px.pie(product_df.head(10), names='Product Category', values='Total Value', hole=0.3, color_discrete_sequence=px.colors.sequential.Greens_r), use_container_width=True)
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
