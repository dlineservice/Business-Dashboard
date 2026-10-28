import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import base64
import os

st.set_page_config(page_title="Executive Dashboard", layout="wide")
C1 = "#2E7D32"
C2 = "#C8E6C9"

st.markdown('<style>.block-container { padding-top: 40px !important; } .kpi-box { background-color: #C8E6C9; padding: 20px; border-radius: 10px; border-left: 5px solid #2E7D32; text-align: center; margin-bottom: 10px; }</style>', unsafe_allow_html=True)

@st.cache_data
def load_all_data():
    try:
        xl = pd.ExcelFile("sales_data.xlsx")
        sn = xl.sheet_names
        sd = {}
        for name in sn:
            sd[name] = pd.read_excel(xl, sheet_name=name)
        return sd, sn
    except Exception as e:
        st.error(f"Excel Loading Error: {e}")
        return {}, []

all_sheets, sheet_names = load_all_data()

if all_sheets:
    mode = st.sidebar.radio("View:", ["Summary", "Sheet View"])
    
    if mode == "Summary":
        total_sales = 0.0
        total_qty = 0.0
        total_overdue = 0.0
        br = {}
        pr = {}
        aging_operating_df = pd.DataFrame()
        
        for s in all_sheets:
            raw_df = all_sheets.get(s)
            if raw_df is None or raw_df.empty or len(raw_df) <= 1:
                continue
                
            df = raw_df.copy()
            df.columns = [str(c).strip().lower() for c in df.columns]
            
            sc = [c for c in df.columns if "sales" in c or "value" in c or "amount" in c]
            qc = [c for c in df.columns if "qty" in c or "quantity" in c or "count" in c]
            
            if s.strip().lower() == "salessummary" and len(sc) > 0:
                try:
                    s_col = sc[0]
                    v = pd.to_numeric(df[s_col], errors="coerce").dropna()
                    if not v.empty:
                        total_sales = float(v.values[-1])
                    bc = [c for c in df.columns if "branch" in c or "zone" in c or "location" in c or "state" in c]
                    if len(bc) > 0:
                        b_col = bc[0]
                        for idx, row in df.iterrows():
                            n = str(row[b_col]).upper().strip()
                            v_br = pd.to_numeric(row[s_col], errors="coerce")
                            if pd.notna(v_br) and n not in ["NAN", "", "GRAND TOTAL", "TOTAL"]:
                                br[n] = br.get(n, 0.0) + float(v_br)
                except:
                    pass
            
            if s.strip().lower() == "dailysalesreport" and len(qc) > 0:
                try:
                    q_col = qc[0]
                    v = pd.to_numeric(df[q_col], errors="coerce").dropna()
                    if not v.empty:
                        total_qty = float(v.values[-1])
                except:
                    pass
                    
            if s.strip().replace(" ", "").lower() == "agingsummary_operatingbranch":
                try:
                    h, w = raw_df.shape
                    if h >= 30 and w >= 10:
                        v_cell = raw_df.iloc[29, 9]
                        v_num = pd.to_numeric(v_cell, errors="coerce")
                        if pd.notna(v_num):
                            total_overdue = float(v_num)
                            
                    r_df = raw_df.iloc[1:].copy()
                    r_df.columns = [str(c).strip() for c in raw_df.columns]
                    
                    r_map = {}
                    for c in r_df.columns:
                        if "branch" in str(c).lower():
                            r_map[c] = "Branch"
                        elif "not over" in str(c).lower():
                            r_map[c] = "Not Over Due"
                        elif "amount" in str(c).lower() or "value" in str(c).lower() or "balance" in str(c).lower() or "overdue" in str(c).lower():
                            r_map[c] = "Amount"
                        else:
                            r_map[c] = c
                            
                    aging_operating_df = r_df.rename(columns=r_map)
                    if "Branch" in aging_operating_df.columns:
                        aging_operating_df = aging_operating_df[aging_operating_df["Branch"].notna()]
                        aging_operating_df = aging_operating_df[~aging_operating_df["Branch"].astype(str).str.contains("Total|GRAND", case=False, na=False)]
                        for c in ["Amount", "Not Over Due"]:
                            if c in aging_operating_df.columns:
                                aging_operating_df[c] = pd.to_numeric(aging_operating_df[c], errors="coerce").fillna(0)
                except:
                    pass
                    
            if s.strip().lower() == "salespdtsummary" and len(sc) > 0:
                try:
                    pc = [c for c in df.columns if "product" in c or "item" in c or "category" in c or "fly" in c]
                    if len(pc) > 0:
                        p_col = pc[0]
                        s_col = sc[0]
                        for idx, row in df.iterrows():
                            n = str(row[p_col]).strip()
                            v_p = pd.to_numeric(row[s_col], errors="coerce")
                            if pd.notna(v_p) and n not in ["nan", "", "GRAND TOTAL", "TOTAL"]:
                                pr[n] = pr.get(n, 0.0) + float(v_p)
                except:
                    pass

        branch_df = pd.DataFrame(list(br.items()), columns=["Branch", "Net Sales"]).sort_values(by="Net Sales", ascending=False)
        product_df = pd.DataFrame(list(pr.items()), columns=["Product Category", "Total Value"]).sort_values(by="Total Value", ascending=False)

        if os.path.exists("FT_Logo.png"):
            with open("FT_Logo.png", "rb") as f:
                encoded_string = base64.b64encode(f.read()).decode()
            st.markdown(f'<div style="display: flex; justify-content: center;"><img src="data:image/png;base64,{encoded_string}" style="max-width: 280px;"></div>', unsafe_allow_html=True)
            
        st.markdown("<h1 style='text-align: center; color: #2E7D32; margin-top: -10px; margin-bottom: 15px;'>Monthly Performance Dashboard</h1>", unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        col1.markdown(f"<div class='kpi-box'><h3>Total Sales</h3><h2>₹ {total_sales:,.2f}</h2></div>", unsafe_allow_html=True)
        col2.markdown(f"<div class='kpi-box'><h3>Total Qty</h3><h2>{total_qty:,.0f} Pcs</h2></div>", unsafe_allow_html=True)
        col3.markdown(f"<div class='kpi-box'><h3>Overdue Amount</h3><h2>₹ {total_overdue:,.2f}</h2></div>", unsafe_allow_html=True)
        
        g1, g2 = st.columns(2)
        if not branch_df.empty:
            g1.plotly_chart(px.bar(branch_df, x="Branch", y="Net Sales", text_auto=".2s", color_discrete_sequence=[C1], template="plotly_white", title="<b>Net Sales by Branch</b>"), use_container_width=True)
        if not product_df.empty:
            g2.plotly_chart(px.pie(product_df.head(10), names="Product Category", values="Total Value", hole=0.3, color_discrete_sequence=px.colors.sequential.Greens_r, title="<b>Sales Share by Product Category (Top 10)</b>"), use_container_width=True)
            
        if not aging_operating_df.empty and "Branch" in aging_operating_df.columns and "Amount" in aging_operating_df.columns:
            st.markdown("<hr>", unsafe_allow_html=True)
            plot_df = aging_operating_df.sort_values(by="Amount", ascending=False)
            fig_combo = go.Figure()
            fig_combo.add_trace(go.Bar(x=plot_df["Branch"], y=plot_df["Amount"], name="Overdue Amount", marker_color="#E53935"))
            if "Not Over Due" in plot_df.columns:
                fig_combo.add_trace(go.Scatter(x=plot_df["Branch"], y=plot_df["Not Over Due"], name="Not Over Due", mode="lines+markers", line=dict(color="#1E88E5", width=3)))
            fig_combo.update_layout(title="<b>Customer Overdue Summary</b>", template="plotly_white", barmode="group")
            st.plotly_chart(fig_combo, use_container_width=True)
            
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown("<h2 style='text-align: center; color: #2E7D32;'>Stock Aging Summary</h2>", unsafe_allow_html=True)
        
        target_sheet = [s for s in all_sheets if "aging" in s.lower() and "stock" in s.lower()]
        sheet_key = target_sheet[0] if len(target_sheet) > 0 else ("StockAgingSummary" if "StockAgingSummary" in all_sheets else None)
        
        if sheet_key:
            aging_df = all_sheets.get(sheet_key).copy()
            aging_df.columns = [str(c).strip() for c in aging_df.columns]
            
            rename_map = {}
            for c in aging_df.columns:
                if "branch" in str(c).lower():
                    rename_map[c] = "Branch"
                elif "age" in str(c).lower():
                    rename_map[c] = "AgeDays"
                elif "pdt" in str(c).lower() or "fly" in str(c).lower():
                    rename_map[c] = "PDT Fly"
                else:
                    rename_map[c] = c
                    
            aging_df = aging_df.rename(columns=rename_map)
            
            val_cols = []
            for c in aging_df.columns:
                is_num = aging_df[c].dtype in ["int64", "float64"]
                has_keyword = any(k in str(c).lower() for k in ["qty", "val", "amount", "balance", "pcs"])
                if (is_num or has_keyword) and c not in ["Branch", "AgeDays", "PDT Fly"]:
                    val_cols.append(c)
            
            if "Branch" in aging_df.columns and len(val_cols) > 0:
                clean_aging = aging_df[aging_df["Branch"].notna()].copy()
                clean_aging = clean_aging[~clean_aging["Branch"].astype(str).str.contains("Total|GRAND", case=False, na=False)]
                for vc in val_cols:
                    clean_aging[vc] = pd.to_numeric(clean_aging[vc], errors="coerce").fillna(0)
                summary_aging = clean_aging.groupby("Branch")[val_cols].sum().reset_index().sort_values(by=val_cols, ascending=False)
                fig_age = px.bar(summary_aging, x="Branch", y=val_cols, title="<b>Stock Breakdown by Branch</b>", barmode="group", template="plotly_white")
                st.plotly_chart(fig_age, use_container_width=True)
            else:
                st.dataframe(aging_df.head(10), use_container_width=True)
        else:
            st.warning("Stock Aging Summary sheet could not be located in the current workbook.")
            
    elif mode == "Sheet View":
        st.markdown("<h2 style='color: #2E7D32; text-align: center;'>Workbook Raw Sheet Explorer</h2>", unsafe_allow_html=True)
        selected_sheet = st.selectbox("Choose a raw datasheet to examine:", sheet_names)
        if selected_sheet in all_sheets:
            st.markdown(f"### Visualizing: {selected_sheet}")
            st.dataframe(all_sheets.get(selected_sheet), use_container_width=True)
        else:
            st.error("Selected sheet data is unavailable.")
else:
    st.error("No data found or sales_data.xlsx missing.")
