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
from tc_reconciliation import (
    run_full_reconciliation,
    generate_reconciliation_excel,
)

# ── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="FYW Pending Order Report",
    page_icon="📦",
    layout="wide",
)

st.title("📦 FYW Pending Order Report")
st.markdown("Upload your daily FYW dashboard export and marketplace files to generate reports.")

tab_report, tab_recon = st.tabs(["📊 Pending Order Report", "🔄 TC Reconciliation Check"])

# ════════════════════════════════════════════════════════════════════════════
# SIDEBAR — Shared file uploads
# ════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.header("📁 Upload Files")

    fyw_csv = st.file_uploader(
        "FYW Dashboard CSV (ALL-DD-Mon-YYYY.csv) *",
        type=["csv"],
        help="Required — the main FYW order export (this is 'TC')"
    )

    st.markdown("---")
    st.markdown("**Shopee Seller Centre Files** _(per brand)_")
    melissa_file = st.file_uploader("Melissa Shopee Order (.xlsx)", type=["xlsx"], key="melissa_shopee")
    ipanema_file = st.file_uploader("Ipanema Shopee Order (.xlsx)", type=["xlsx"], key="ipanema_shopee")
    cspace_file  = st.file_uploader("CSpace Shopee Order (.xlsx)",  type=["xlsx"], key="cspace_shopee")

    st.markdown("---")
    st.markdown("**Other Marketplace Files** _(for TC reconciliation)_")
    tiktok_file = st.file_uploader("TikTok Order Export (.xlsx/.csv)", type=["xlsx", "csv"], key="tiktok_file")

    st.markdown("**Lazada** _(per brand)_")
    lazada_melissa_file = st.file_uploader("Lazada - Melissa Order Export (.xlsx/.csv)", type=["xlsx", "csv"], key="lazada_melissa")
    lazada_ipanema_file = st.file_uploader("Lazada - Ipanema Order Export (.xlsx/.csv)", type=["xlsx", "csv"], key="lazada_ipanema")
    lazada_cspace_file  = st.file_uploader("Lazada - CSpace Order Export (.xlsx/.csv)",  type=["xlsx", "csv"], key="lazada_cspace")

    zalora_file = st.file_uploader("Zalora Order Export (.xlsx/.csv)", type=["xlsx", "csv"], key="zalora_file")

    st.markdown("---")
    report_date = st.date_input("Report Date", value=datetime.today())

if not fyw_csv:
    st.info("👈 Upload the FYW CSV file to get started.")
    st.stop()

# ════════════════════════════════════════════════════════════════════════════
# TAB 1 — PENDING ORDER REPORT
# ════════════════════════════════════════════════════════════════════════════
with tab_report:
    with st.spinner("Loading files..."):
        try:
            pending = load_fyw_csv(fyw_csv)
        except Exception as e:
            st.error(f"❌ Failed to load FYW CSV: {e}")
            st.stop()

        shopee_map = {'MELISSA': melissa_file, 'IPANEMA': ipanema_file, 'CSPACE': cspace_file}
        mp_sla_map, shopee_tracking = load_shopee_files(shopee_map)
        df_report = build_report_rows(pending, mp_sla_map, shopee_tracking)

    st.subheader("📊 Summary")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Pending Items",  len(df_report))
    col2.metric("Unique Orders",        df_report['Order ID'].nunique())
    col3.metric("Total Value (MYR)",    f"{df_report['Sold Price (MYR)'].sum():,.2f}")
    col4.metric("Active Channels",      df_report['Channel'].nunique())

    st.subheader("🗂️ Pivot: Orders by Nickname × FYW SLA")
    pivot_raw = df_report.pivot_table(
        index='Nickname', columns='FYW SLA', values='Order ID',
        aggfunc='count', fill_value=0
    )
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

    col_a, col_b, col_c, col_d = st.columns(4)
    col_a.markdown("🔴 **Past due**")
    col_b.markdown("🟠 **Yesterday**")
    col_c.markdown("🟢 **Today**")
    col_d.markdown("🔵 **Grand Total**")

    st.subheader("🏷️ Brand Summary")
    brand_summary = df_report.groupby('Brand').agg(
        Orders=('Order ID', 'nunique'),
        Items=('Qty', 'sum'),
        Value=('Sold Price (MYR)', 'sum'),
        Channels=('Channel', lambda x: ', '.join(sorted(set(x))))
    ).reset_index()
    brand_summary['Value'] = brand_summary['Value'].apply(lambda x: f"MYR {x:,.2f}")
    st.dataframe(brand_summary, use_container_width=True, hide_index=True)

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
        key="download_pending_report",
    )
    st.success(f"✅ **{filename}** — {len(df_report)} items | MYR {df_report['Sold Price (MYR)'].sum():,.2f}")

# ════════════════════════════════════════════════════════════════════════════
# TAB 2 — TC RECONCILIATION CHECK
# ════════════════════════════════════════════════════════════════════════════
with tab_recon:
    st.markdown(
        "Checks whether orders from each marketplace export have successfully been "
        "**pushed to TC** (i.e. appear in the FYW dashboard CSV). "
        "Upload marketplace files in the sidebar to run this check."
    )

    mp_files = {
        'Shopee - Melissa': melissa_file,
        'Shopee - Ipanema': ipanema_file,
        'Shopee - CSpace':  cspace_file,
        'TikTok':           tiktok_file,
        'Lazada - Melissa': lazada_melissa_file,
        'Lazada - Ipanema': lazada_ipanema_file,
        'Lazada - CSpace':  lazada_cspace_file,
        'Zalora':           zalora_file,
    }
    uploaded_mp_files = {k: v for k, v in mp_files.items() if v is not None}

    if not uploaded_mp_files:
        st.info("👈 Upload at least one marketplace file (Shopee, TikTok, Lazada, or Zalora) in the sidebar to run the reconciliation check.")
        st.stop()

    with st.spinner("Loading FYW CSV for reconciliation..."):
        fyw_csv.seek(0)
        fyw_raw_df = pd.read_csv(fyw_csv)

    with st.spinner("Cross-checking marketplace orders against TC..."):
        try:
            results = run_full_reconciliation(fyw_raw_df, uploaded_mp_files)
        except Exception as e:
            st.error(f"❌ Reconciliation failed: {e}")
            st.stop()

    st.subheader("📊 Reconciliation Summary")

    total_missing = sum(r.get('missing_count', 0) for r in results.values())
    total_mp_orders = sum(r.get('total_mp_orders', 0) for r in results.values())

    col1, col2, col3 = st.columns(3)
    col1.metric("Marketplaces Checked", len(results))
    col2.metric("Total MP Orders",      total_mp_orders)
    col3.metric(
        "Total Missing from TC",
        total_missing,
        delta=f"{total_missing} not pushed" if total_missing > 0 else "All synced ✅",
        delta_color="inverse" if total_missing > 0 else "normal",
    )

    summary_rows = []
    for label, res in results.items():
        if 'error' in res:
            summary_rows.append({
                'Marketplace': label,
                'Status': f"⚠️ {res['error']}",
                'Total MP Orders': '-',
                'Matched in TC': '-',
                'Missing from TC': '-',
                'Match Rate': '-',
            })
        else:
            status = "✅ All synced" if res['missing_count'] == 0 else f"🔴 {res['missing_count']} missing"
            summary_rows.append({
                'Marketplace': label,
                'Status': status,
                'Total MP Orders': res['total_mp_orders'],
                'Matched in TC': res['matched_count'],
                'Missing from TC': res['missing_count'],
                'Match Rate': f"{res['match_rate']}%" if res['match_rate'] is not None else '-',
            })

    summary_df = pd.DataFrame(summary_rows)

    def highlight_status(row):
        if '🔴' in str(row['Status']):
            return ['background-color: #FFCCCC'] * len(row)
        elif '✅' in str(row['Status']):
            return ['background-color: #C6EFCE'] * len(row)
        elif '⚠️' in str(row['Status']):
            return ['background-color: #FFE699'] * len(row)
        return [''] * len(row)

    st.dataframe(
        summary_df.style.apply(highlight_status, axis=1),
        use_container_width=True,
        hide_index=True,
    )

    # ── Missing Orders Detail ────────────────────────────────────────────────
    st.subheader("🔍 Missing Order Details")

    has_missing = any(r.get('missing_count', 0) > 0 for r in results.values())
    if not has_missing:
        st.success("🎉 All marketplace orders have been successfully pushed to TC!")
    else:
        for label, res in results.items():
            if res.get('missing_count', 0) > 0:
                with st.expander(f"🔴 {label} — {res['missing_count']} orders missing from TC"):
                    missing_df = pd.DataFrame({
                        '#': range(1, len(res['missing_ids']) + 1),
                        'Order ID / Order Number': res['missing_ids'],
                    })
                    st.dataframe(missing_df, use_container_width=True, hide_index=True)

    # ── Download ──────────────────────────────────────────────────────────────
    st.subheader("⬇️ Download Reconciliation Report")
    with st.spinner("Generating reconciliation Excel..."):
        recon_buffer = generate_reconciliation_excel(
            results,
            report_date=datetime.combine(report_date, datetime.min.time())
        )

    recon_filename = f"TC_Reconciliation_{report_date.strftime('%d%b%Y')}.xlsx"
    st.download_button(
        label="📥 Download Reconciliation Excel",
        data=recon_buffer,
        file_name=recon_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        key="download_reconciliation_report",
    )
