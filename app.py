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
    period = st.session_state.current_period
    metrics = data["metrics"]
    df_detail = data["df_detail"]
    df_rekap = data["df_rekap"]
    top_5 = data["top_5"]

    st.markdown("---")

    # 1. KPI Metric Cards
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Total Beban OPEX</div>
            <div class="metric-value">{format_rupiah(metrics['total_saldo'])}</div>
            <div class="metric-badge badge-primary">Periode {period}</div>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Pos Akun Unik</div>
            <div class="metric-value">{metrics['total_accounts']} <span style="font-size:0.9rem; font-weight:500; color:#64748B;">Akun</span></div>
            <div class="metric-badge badge-success">100% Dinamis</div>
        </div>
        """, unsafe_allow_html=True)

    with c3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Total Transaksi</div>
            <div class="metric-value">{metrics['total_transactions']} <span style="font-size:0.9rem; font-weight:500; color:#64748B;">Baris</span></div>
            <div class="metric-badge badge-warning">Lengkap</div>
        </div>
        """, unsafe_allow_html=True)

    with c4:
        reversal_txt = f"{metrics['reversal_count']} Reversal"
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Koreksi / Saldo Minus</div>
            <div class="metric-value" style="color: {'#DC2626' if metrics['reversal_count'] > 0 else '#059669'};">
                {format_rupiah(metrics['reversal_amount'])}
            </div>
            <div class="metric-badge badge-danger">{reversal_txt}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

    # 2. Visualization: Top 5 Highest Expense Accounts
    st.subheader("📊 Top 5 Beban Pengeluaran Terbesar")
    
    if not top_5.empty:
        # Prepare chart dataframe
        chart_df = top_5.copy()
        # Sort ascending for horizontal bar chart
        chart_df = chart_df.sort_values(by="Total Saldo", ascending=True)
        chart_df["Saldo_Formatted"] = chart_df["Total Saldo"].apply(format_rupiah)

        chart = alt.Chart(chart_df).mark_bar(
            cornerRadiusTopRight=6,
            cornerRadiusBottomRight=6,
            color="#1E3A8A"
        ).encode(
            x=alt.X('Total Saldo:Q', title='Total Beban (IDR)', axis=alt.Axis(format=',.0f')),
            y=alt.Y('Account Name:N', sort=None, title='Nama Akun OPEX'),
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

    # 3. Interactive Data Previews
    tab_rekap, tab_detail, tab_minus = st.tabs([
        f"📋 Rekapitulasi Pos Akun ({len(df_rekap)})",
        f"📄 Detail Transaksi ({len(df_detail)})",
        f"🔍 Audit Saldo Minus / Koreksi ({metrics['reversal_count']})"
    ])

    with tab_rekap:
        st.caption("Ringkasan saldo bersih per pos akun (telah memperhitungkan transaksi koreksi/minus):")
        df_rekap_display = df_rekap.copy()
        total_sum = df_rekap_display["Total Saldo"].sum()
        df_rekap_display["% Proporsi"] = (df_rekap_display["Total Saldo"] / total_sum * 100).map("{:.2f}%".format) if total_sum != 0 else "0.00%"
        df_rekap_display["Total Saldo"] = df_rekap_display["Total Saldo"].apply(format_rupiah)
        st.dataframe(df_rekap_display, use_container_width=True, hide_index=True)

    with tab_detail:
        col_search, col_acc = st.columns([2, 2])
        with col_search:
            search_query = st.text_input("Cari keterangan atau batch transaksi:", placeholder="Ketik kata kunci...")
        with col_acc:
            account_list = ["Semua Akun"] + sorted(df_detail["Account Name"].unique().tolist())
            selected_acc = st.selectbox("Filter berdasarkan Akun:", account_list)

        filtered_detail = df_detail.copy()
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
        st.caption("Daftar transaksi pembalik (reversal/credit) atau penyesuaian yang bernilai minus:")
        df_minus = df_detail[df_detail["Saldo"] < 0].copy()
        if not df_minus.empty:
            df_minus_display = df_minus.copy()
            df_minus_display["Saldo"] = df_minus_display["Saldo"].apply(format_rupiah)
            st.dataframe(df_minus_display, use_container_width=True, hide_index=True)
        else:
            st.info("Tidak ada transaksi bertanda minus pada dokumen ini.")

    # 4. Download Section (2 Official File Outputs)
    st.markdown("---")
    st.subheader("📥 Unduh Berkas Pembukuan Resmi")
    st.caption("Unduh berkas resmi yang telah diformat sesuai standar pelaporan dan audit FAB:")

    # Generate output files in memory
    with st.spinner("Menyiapkan berkas unduhan..."):
        excel_buffer = generate_detail_excel(df_detail, df_rekap, period)
        pdf_buffer = generate_rekap_pdf(df_rekap, period, metrics["total_saldo"])

    excel_file_name = f"Detail Realisasi OPEX - {period}.xlsx"
    pdf_file_name = f"Rekap Realisasi OPEX - {period}.pdf"

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
        st.caption("Format PDF formal siap cetak dengan tanda tangan & baris TOTAL SEMUA.")

    with col_dl_excel:
        st.download_button(
            label=f"📊 Download Detail Excel ({excel_file_name})",
            data=excel_buffer,
            file_name=excel_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.caption("Format spreadsheet standar akuntansi (6 kolom baris 1, pemisah ribuan).")

else:
    # Empty State Guidance
    st.info("👋 Silakan pilih atau tarik dokumen laporan realisasi OPEX (format PDF atau teks mentah Oracle Financials/IBS) di atas untuk memulai rekonsiliasi.")
