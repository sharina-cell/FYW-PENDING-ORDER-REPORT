"""
FYW Pending Order Report — Streamlit App
Run with: streamlit run app.py
"""

import streamlit as st
import pandas as pd
from datetime import datetime
from report_engine import (
    load_fyw_csv,
    load_shopee_files,
    build_report_rows,
    generate_excel,
)

# ── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="FYW Pending Order Report",
    page_icon="📦",
    layout="wide",
)

st.title("📦 FYW Pending Order Report")
st.markdown("Upload your daily FYW dashboard export and Shopee files to generate the report.")

# ── Sidebar: File Uploads ─────────────────────────────────────────────────────
with st.sidebar:
    st.header("📁 Upload Files")

    fyw_csv = st.file_uploader(
        "FYW Dashboard CSV (ALL-DD-Mon-YYYY.csv) *",
        type=["csv"],
        help="Required — the main FYW order export"
    )

    st.markdown("---")
    st.markdown("**Shopee Seller Centre Files** _(optional, for MP SLA lookup)_")
    melissa_file = st.file_uploader("Melissa Shopee Order (.xlsx)", type=["xlsx"])
    ipanema_file = st.file_uploader("Ipanema Shopee Order (.xlsx)", type=["xlsx"])
    cspace_file  = st.file_uploader("CSpace Shopee Order (.xlsx)",  type=["xlsx"])

    st.markdown("---")
    report_date = st.date_input("Report Date", value=datetime.today())

# ── Main Panel ────────────────────────────────────────────────────────────────
if not fyw_csv:
    st.info("👈 Upload the FYW CSV file to get started.")
    st.stop()

with st.spinner("Loading files..."):
    try:
        pending = load_fyw_csv(fyw_csv)
    except Exception as e:
        st.error(f"❌ Failed to load FYW CSV: {e}")
        st.stop()

    shopee_map = {'MELISSA': melissa_file, 'IPANEMA': ipanema_file, 'CSPACE': cspace_file}
    mp_sla_map, shopee_tracking = load_shopee_files(shopee_map)
    df_report = build_report_rows(pending, mp_sla_map, shopee_tracking)

# ── Summary Metrics ───────────────────────────────────────────────────────────
st.subheader("📊 Summary")
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total Pending Items",  len(df_report))
col2.metric("Unique Orders",        df_report['Order ID'].nunique())
col3.metric("Total Value (MYR)",    f"{df_report['Sold Price (MYR)'].sum():,.2f}")
col4.metric("Active Channels",      df_report['Channel'].nunique())

# ── Pivot Table ───────────────────────────────────────────────────────────────
st.subheader("🗂️ Pivot: Orders by Nickname × FYW SLA")

pivot_raw = df_report.pivot_table(
    index='Nickname', columns='FYW SLA', values='Order ID',
    aggfunc='count', fill_value=0
)
# Sort SLA date columns chronologically (DD/MM/YYYY) so columns display correctly
def _sort_date_col(col_name):
    try:
        return datetime.strptime(col_name, '%d/%m/%Y').date()
    except Exception:
        return datetime.max.date()

sla_cols = sorted([c for c in pivot_raw.columns], key=_sort_date_col)
pivot_raw = pivot_raw.reindex(columns=sla_cols)
pivot_raw['Grand Total'] = pivot_raw.sum(axis=1)
total_row = pivot_raw.sum(axis=0)
total_row.name = 'Grand Total'
pivot_display = pd.concat([pivot_raw, total_row.to_frame().T]).astype(int)
pivot_display = pivot_display.replace(0, '-')

def style_pivot(df):
    styles = pd.DataFrame('', index=df.index, columns=df.columns)
    today  = datetime.today().date()
    yest   = (datetime.today() - pd.Timedelta(days=1)).date()
    for col in df.columns:
        if col == 'Grand Total':
            styles[col] = 'background-color: #DDEEFF; font-weight: bold'
        else:
            try:
                col_date = datetime.strptime(col, '%d/%m/%Y').date()
                if col_date < yest:
                    styles[col] = 'background-color: #FFCCCC; color: #9C0006; font-weight: bold'
                elif col_date == yest:
                    styles[col] = 'background-color: #FCE4D6; color: #833C00; font-weight: bold'
                elif col_date == today:
                    styles[col] = 'background-color: #C6EFCE; color: #006100; font-weight: bold'
            except Exception:
                pass
    return styles

st.dataframe(pivot_display.style.apply(style_pivot, axis=None), use_container_width=True)

# Legend
col_a, col_b, col_c, col_d = st.columns(4)
col_a.markdown("🔴 **Past due**")
col_b.markdown("🟠 **Yesterday**")
col_c.markdown("🟢 **Today**")
col_d.markdown("🔵 **Grand Total**")

# ── Brand Summary ─────────────────────────────────────────────────────────────
st.subheader("🏷️ Brand Summary")
brand_summary = df_report.groupby('Brand').agg(
    Orders=('Order ID', 'nunique'),
    Items=('Qty', 'sum'),
    Value=('Sold Price (MYR)', 'sum'),
    Channels=('Channel', lambda x: ', '.join(sorted(set(x))))
).reset_index()
brand_summary['Value'] = brand_summary['Value'].apply(lambda x: f"MYR {x:,.2f}")
st.dataframe(brand_summary, use_container_width=True, hide_index=True)

# ── Detailed Orders ───────────────────────────────────────────────────────────
st.subheader("📋 All Pending Orders")

with st.expander("🔍 Filter options"):
    c1, c2, c3 = st.columns(3)
    brand_filter   = c1.multiselect("Brand",   options=sorted(df_report['Brand'].unique()),   default=list(df_report['Brand'].unique()))
    channel_filter = c2.multiselect("Channel", options=sorted(df_report['Channel'].unique()), default=list(df_report['Channel'].unique()))
    status_filter  = c3.multiselect("Status",  options=sorted(df_report['FYW Status'].unique()), default=list(df_report['FYW Status'].unique()))

filtered = df_report[
    df_report['Brand'].isin(brand_filter) &
    df_report['Channel'].isin(channel_filter) &
    df_report['FYW Status'].isin(status_filter)
]
st.dataframe(filtered.drop(columns=['Nickname']), use_container_width=True, hide_index=True)
st.caption(f"Showing {len(filtered)} of {len(df_report)} orders")

# ── Download ──────────────────────────────────────────────────────────────────
st.subheader("⬇️ Download Excel Report")
with st.spinner("Generating Excel..."):
    excel_buffer = generate_excel(
        df_report,
        report_date=datetime.combine(report_date, datetime.min.time())
    )

filename = f"FYW_Pending_Orders_{report_date.strftime('%d%b%Y')}.xlsx"
st.download_button(
    label="📥 Download Excel Report",
    data=excel_buffer,
    file_name=filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
st.success(f"✅ **{filename}** — {len(df_report)} items | MYR {df_report['Sold Price (MYR)'].sum():,.2f}")
