"""
FORMS - FAB OPEX Reconciliation & Monitoring System
Web Application: Mobile-first Streamlit interface for automated reconciliation,
data processing, and formal Excel & PDF report generation.
"""

import io
import os
import sys
from pathlib import Path
import streamlit as st
import pandas as pd
import altair as alt

# Ensure repo root is on sys.path for robust cloud imports
ROOT_DIR = Path(__file__).resolve().parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from parser_engine import process_opex_file
from excel_generator import generate_detail_excel
from pdf_generator import generate_rekap_pdf, format_rupiah


# Page Configuration
st.set_page_config(
    page_title="FORMS | FAB OPEX Reconciliation & Monitoring System",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for Executive Financial Styling & Mobile Optimization
st.markdown("""
<style>
    /* Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    /* Main Container Padding */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3rem;
        padding-left: 1.25rem;
        padding-right: 1.25rem;
        max-width: 1200px;
    }

    /* Header Banner */
    .hero-banner {
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 50%, #2563EB 100%);
        color: white;
        padding: 1.75rem 1.5rem;
        border-radius: 16px;
        margin-bottom: 1.5rem;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.2);
    }
    .hero-title {
        font-size: 1.85rem;
        font-weight: 800;
        letter-spacing: -0.02em;
        margin-bottom: 0.25rem;
        color: #FFFFFF;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    .hero-subtitle {
        font-size: 0.95rem;
        font-weight: 500;
        color: #93C5FD;
        margin-bottom: 0.2rem;
    }
    .hero-tagline {
        font-size: 0.85rem;
        font-style: italic;
        color: #E2E8F0;
        opacity: 0.9;
    }

    /* KPI Cards */
    .metric-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 14px;
        padding: 1.15rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        transition: transform 0.15s ease, box-shadow 0.15s ease;
        height: 100%;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 12px rgba(0,0,0,0.05);
    }
    .metric-label {
        font-size: 0.78rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #64748B;
        margin-bottom: 0.4rem;
    }
    .metric-value {
        font-size: 1.45rem;
        font-weight: 800;
        color: #0F172A;
        line-height: 1.2;
    }
    .metric-badge {
        display: inline-block;
        font-size: 0.75rem;
        font-weight: 600;
        padding: 0.2rem 0.55rem;
        border-radius: 9999px;
        margin-top: 0.4rem;
    }
    .badge-primary { background: #EFF6FF; color: #1D4ED8; }
    .badge-success { background: #ECFDF5; color: #047857; }
    .badge-danger { background: #FEF2F2; color: #B91C1C; }
    .badge-warning { background: #FFFBEB; color: #B45309; }

    /* Action Buttons */
    .stButton > button {
        border-radius: 10px;
        font-weight: 600;
        padding: 0.55rem 1.25rem;
        transition: all 0.2s ease;
    }
    .stDownloadButton > button {
        width: 100%;
        border-radius: 12px;
        font-weight: 700;
        padding: 0.75rem 1rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }

    /* Mobile specific adjustments */
    @media (max-width: 768px) {
        .block-container {
            padding-left: 0.75rem;
            padding-right: 0.75rem;
            padding-top: 1rem;
        }
        .hero-title {
            font-size: 1.35rem;
        }
        .hero-subtitle {
            font-size: 0.85rem;
        }
        .metric-value {
            font-size: 1.25rem;
        }
    }
</style>
""", unsafe_allow_html=True)


# Application Header Banner
st.markdown("""
<div class="hero-banner">
    <div class="hero-title">
        <span>💼</span> FORMS
    </div>
    <div class="hero-subtitle">FAB OPEX Reconciliation & Monitoring System</div>
    <div class="hero-tagline">"Fast, Accurate, and Accessible OPEX Insights."</div>
</div>
""", unsafe_allow_html=True)


# Session State Initialization
if "processed_data" not in st.session_state:
    st.session_state.processed_data = None
if "current_period" not in st.session_state:
    st.session_state.current_period = "Aug-26"


# Top Action Section: File Upload (100% Real Data)
uploaded_file = st.file_uploader(
    "Pilih atau Tarik Berkas Laporan Realisasi OPEX (PDF / Teks Mentah):",
    type=["pdf", "txt", "prn", "rep", "csv"],
    help="Mendukung laporan PDF dan berkas teks mentah keluaran Oracle Financials / IBS."
)

# Trigger Process Button if file is uploaded
if uploaded_file is not None:
    file_size_kb = len(uploaded_file.getvalue()) / 1024
    st.caption(f"📁 Berkas terpilih: **{uploaded_file.name}** ({file_size_kb:.1f} KB)")
    
    btn_process = st.button("🚀 Proses Data Laporan", type="primary", use_container_width=True)
    if btn_process or st.session_state.processed_data is None:
        with st.spinner("Membaca dan merekonsiliasi transaksi OPEX..."):
            file_bytes = uploaded_file.getvalue()
            res = process_opex_file(file_bytes, uploaded_file.name)
            if res["success"]:
                st.session_state.processed_data = res
                st.session_state.current_period = res["period"]
                st.success(f"Berhasil merekonsiliasi dokumen! Periode terdeteksi: **{res['period']}**")
            else:
                st.error(f"Peringatan: {res.get('error', 'Format tidak dikenali.')}")
else:
    # Reset session state when no file is uploaded
    st.session_state.processed_data = None


# Render Processed Dashboard
if st.session_state.processed_data is not None:
    data = st.session_state.processed_data
    detected_period = st.session_state.current_period
    df_full = data["df_detail"]

    # 1. Date & Period Extraction for Filtering (YTD vs Per-Bulan)
    df_work = df_full.copy()
    parsed_dt = pd.to_datetime(df_work["Trans. Date"], dayfirst=True, errors="coerce")
    df_work["_dt"] = parsed_dt
    df_work["_period"] = parsed_dt.dt.to_period("M")

    # Discover all distinct month periods present in the uploaded document
    valid_periods = [p for p in df_work["_period"].dropna().unique()]
    valid_periods.sort()

    month_names_id = {
        1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
        7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember"
    }

    period_options = [f"📊 Semua Periode (YTD {detected_period})"]
    period_map = {f"📊 Semua Periode (YTD {detected_period})": "YTD"}

    for p in valid_periods:
        m_id = month_names_id.get(p.month, p.strftime('%B'))
        label = f"📅 {m_id} {p.year} ({p.strftime('%b-%y')})"
        period_options.append(label)
        period_map[label] = p

    st.markdown("---")

    # Period Dropdown Filter
    col_filter, col_badge = st.columns([3, 2], gap="medium")
    with col_filter:
        selected_period_label = st.selectbox(
            "🔎 Filter Cakupan Periode Laporan (YTD / Per Bulan):",
            options=period_options,
            index=0,
            help="Pilih 'Semua Periode (YTD)' untuk data kumulatif, atau pilih bulan tertentu untuk menyaring data dan mengunduh laporan bulan tersebut."
        )

    # Determine Active Dataset based on Selection
    sel_val = period_map[selected_period_label]
    if sel_val == "YTD":
        df_active = df_work.drop(columns=["_dt", "_period"])
        active_period_name = f"YTD {detected_period}"
        active_badge_txt = f"Kumulatif YTD {detected_period}"
    else:
        df_active = df_work[df_work["_period"] == sel_val].drop(columns=["_dt", "_period"])
        m_id = month_names_id.get(sel_val.month, sel_val.strftime('%B'))
        active_period_name = f"{sel_val.strftime('%b-%y')}"
        active_badge_txt = f"Bulan {m_id} {sel_val.year}"

    with col_badge:
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        st.info(f"📌 Menampilkan: **{active_badge_txt}** ({len(df_active)} Transaksi)")

    # Dynamic Recalculation for Active Period
    active_total_saldo = float(df_active["Saldo"].sum())
    active_df_rekap = df_active.groupby(["Code Account", "Account Name"], as_index=False)["Saldo"].sum()
    active_df_rekap.rename(columns={"Saldo": "Total Saldo"}, inplace=True)
    active_df_rekap.sort_values(by="Total Saldo", ascending=False, inplace=True)
    active_df_rekap.reset_index(drop=True, inplace=True)

    active_total_accounts = int(active_df_rekap["Code Account"].nunique())
    active_total_transactions = len(df_active)
    active_minus_txs = df_active[df_active["Saldo"] < 0]
    active_reversal_count = len(active_minus_txs)
    active_reversal_amount = float(active_minus_txs["Saldo"].sum())
    active_top_5 = active_df_rekap.head(5).copy()

    # 2. KPI Metric Cards
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Total Beban OPEX</div>
            <div class="metric-value">{format_rupiah(active_total_saldo)}</div>
            <div class="metric-badge badge-primary">{active_badge_txt}</div>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Pos Akun Unik</div>
            <div class="metric-value">{active_total_accounts} <span style="font-size:0.9rem; font-weight:500; color:#64748B;">Akun</span></div>
            <div class="metric-badge badge-success">100% Dinamis</div>
        </div>
        """, unsafe_allow_html=True)

    with c3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Total Transaksi</div>
            <div class="metric-value">{active_total_transactions} <span style="font-size:0.9rem; font-weight:500; color:#64748B;">Baris</span></div>
            <div class="metric-badge badge-warning">Lengkap</div>
        </div>
        """, unsafe_allow_html=True)

    with c4:
        reversal_txt = f"{active_reversal_count} Reversal"
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Koreksi / Saldo Minus</div>
            <div class="metric-value" style="color: {'#DC2626' if active_reversal_count > 0 else '#059669'};">
                {format_rupiah(active_reversal_amount)}
            </div>
            <div class="metric-badge badge-danger">{reversal_txt}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

    # 3. Visualization: Top 5 Highest Expense Accounts (Largest on Top)
    st.subheader("📊 Top 5 Beban Pengeluaran Terbesar")
    
    if not active_top_5.empty:
        # Prepare chart dataframe
        chart_df = active_top_5.copy()
        # Sort descending by Total Saldo so the highest spending is #1 at the top
        chart_df = chart_df.sort_values(by="Total Saldo", ascending=False)
        chart_df["Saldo_Formatted"] = chart_df["Total Saldo"].apply(format_rupiah)

        chart = alt.Chart(chart_df).mark_bar(
            cornerRadiusTopRight=6,
            cornerRadiusBottomRight=6,
            color="#1E3A8A"
        ).encode(
            x=alt.X('Total Saldo:Q', title='Total Beban (IDR)', axis=alt.Axis(format=',.0f')),
            y=alt.Y('Account Name:N', sort='-x', title='Nama Akun OPEX'),
            tooltip=[
                alt.Tooltip('Code Account:N', title='Kode Akun'),
                alt.Tooltip('Account Name:N', title='Nama Akun'),
                alt.Tooltip('Saldo_Formatted:N', title='Total Saldo')
            ]
        ).properties(height=240)

        # Text labels on bars
        text = chart.mark_text(
            align='left',
            baseline='middle',
            dx=5,
            fontSize=11,
            fontWeight='bold',
            color='#1E293B'
        ).encode(
            text='Saldo_Formatted:N'
        )

        st.altair_chart(chart + text, use_container_width=True)

    # 4. Interactive Data Previews
    tab_rekap, tab_detail, tab_minus = st.tabs([
        f"📋 Rekapitulasi Pos Akun ({len(active_df_rekap)})",
        f"📄 Detail Transaksi ({len(df_active)})",
        f"🔍 Audit Saldo Minus / Koreksi ({active_reversal_count})"
    ])

    with tab_rekap:
        st.caption(f"Ringkasan saldo bersih per pos akun untuk periode **{active_badge_txt}**:")
        df_rekap_display = active_df_rekap.copy()
        total_sum = df_rekap_display["Total Saldo"].sum()
        df_rekap_display["% Proporsi"] = (df_rekap_display["Total Saldo"] / total_sum * 100).map("{:.2f}%".format) if total_sum != 0 else "0.00%"
        df_rekap_display["Total Saldo"] = df_rekap_display["Total Saldo"].apply(format_rupiah)
        st.dataframe(df_rekap_display, use_container_width=True, hide_index=True)

    with tab_detail:
        col_search, col_acc = st.columns([2, 2])
        with col_search:
            search_query = st.text_input("Cari keterangan atau batch transaksi:", placeholder="Ketik kata kunci...")
        with col_acc:
            account_list = ["Semua Akun"] + sorted(df_active["Account Name"].unique().tolist())
            selected_acc = st.selectbox("Filter berdasarkan Akun:", account_list)

        filtered_detail = df_active.copy()
        if search_query:
            filtered_detail = filtered_detail[
                filtered_detail["Description"].str.contains(search_query, case=False, na=False) |
                filtered_detail["Batch Name"].str.contains(search_query, case=False, na=False)
            ]
        if selected_acc != "Semua Akun":
            filtered_detail = filtered_detail[filtered_detail["Account Name"] == selected_acc]

        df_detail_display = filtered_detail.copy()
        df_detail_display["Saldo"] = df_detail_display["Saldo"].apply(format_rupiah)
        st.dataframe(df_detail_display, use_container_width=True, hide_index=True)

    with tab_minus:
        st.caption(f"Daftar transaksi pembalik (reversal/credit) atau penyesuaian yang bernilai minus untuk **{active_badge_txt}**:")
        df_minus = df_active[df_active["Saldo"] < 0].copy()
        if not df_minus.empty:
            df_minus_display = df_minus.copy()
            df_minus_display["Saldo"] = df_minus_display["Saldo"].apply(format_rupiah)
            st.dataframe(df_minus_display, use_container_width=True, hide_index=True)
        else:
            st.info(f"Tidak ada transaksi bertanda minus pada periode {active_badge_txt}.")

    # 5. Download Section (2 Official File Outputs for Active Period)
    st.markdown("---")
    st.subheader("📥 Unduh Berkas Pembukuan Resmi")
    st.caption(f"Berkas resmi yang diunduh mencakup data untuk periode: **{active_badge_txt}**")

    # Generate output files in memory
    with st.spinner("Menyiapkan berkas unduhan..."):
        excel_buffer = generate_detail_excel(df_active, active_df_rekap, active_period_name)
        pdf_buffer = generate_rekap_pdf(active_df_rekap, active_period_name, active_total_saldo)

    excel_file_name = f"Detail Realisasi OPEX - {active_period_name}.xlsx"
    pdf_file_name = f"Rekap Realisasi OPEX - {active_period_name}.pdf"

    col_dl_pdf, col_dl_excel = st.columns(2, gap="medium")

    with col_dl_pdf:
        st.download_button(
            label=f"📄 Download Rekap PDF ({pdf_file_name})",
            data=pdf_buffer,
            file_name=pdf_file_name,
            mime="application/pdf",
            type="primary",
            use_container_width=True
        )
        st.caption("Format PDF formal: Divisi SERVICE, TTD Kiri FAB, Kanan SUAH, baris TOTAL SEMUA.")

    with col_dl_excel:
        st.download_button(
            label=f"📊 Download Detail Excel ({excel_file_name})",
            data=excel_buffer,
            file_name=excel_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.caption("Format spreadsheet standar akuntansi (kolom Trans. Date terstandar, pemisah ribuan).")

else:
    # Empty State Guidance
    st.info("👋 Silakan pilih atau tarik dokumen laporan realisasi OPEX (format PDF atau teks mentah Oracle Financials/IBS) di atas untuk memulai rekonsiliasi.")
