# FORMS (FAB OPEX Reconciliation & Monitoring System)

> *"Fast, Accurate, and Accessible OPEX Insights."*

FORMS adalah solusi otomatisasi berbasis **Python & Streamlit** untuk mengonversi dokumen laporan realisasi OPEX mentah (teks / PDF dari Oracle Financials / IBS) menjadi dua output pembukuan formal (Excel detail transaksi & PDF rekapitulasi akun), dilengkapi visualisasi interaktif dan metrik ringkas yang dapat diakses langsung melalui peramban ponsel pintar kapan saja.

---

## 🚀 Fitur Utama

- **100% Dinamis & Zero Hardcode:** Mengekstrak setiap kombinasi unik `Code Account` & `Account Name` secara mandiri dari dokumen bulan berjalan.
- **Deteksi Periode Otomatis:** Menangkap periode bulan laporan (misal: *Aug-26*) untuk penamaan berkas resmi keluaran.
- **Akurasi Saldo Riil & Rekonsiliasi Transaksi Minus:** Memperhitungkan transaksi bertanda minus (`-`, `()`, `CR`) secara presisi hingga digit rupiah terakhir.
- **Dukungan Format Luas:** Memproses file PDF cetak sistem (`FNDWRR`) maupun file teks mentah (`.txt`, `.prn`, `.rep`, `.csv`).
- **Luaran Berkas Resmi:**
  - **Excel Detail:** `Detail Realisasi OPEX - [bulan apa].xlsx` (6 kolom standar di Baris 1, format akuntansi ribuan, sheet rekapitulasi pos akun).
  - **PDF Rekapitulasi:** `Rekap Realisasi OPEX - [bulan apa].pdf` (Tabel formal, baris penutup bergaris ganda akuntansi **TOTAL SEMUA**, kolom tanda tangan pengesahan).
- **Mobile-First Responsive Dashboard:** Kartu ringkasan KPI, grafik Top 5 pengeluaran terbesar, tab audit transaksi minus, dan tombol download instan.

---

## 📁 Struktur Direktori

```
FORMS-MINA/
├── app.py                   # Antarmuka Streamlit (Mobile-first, visualisasi Top 5, KPI cards, download buttons)
├── parser_engine.py         # Mesin ekstraksi dinamis PDF & teks, deteksi periode, normalisasi saldo
├── excel_generator.py       # Pembuat berkas Excel formal (openpyxl) dengan standar pembukuan FAB
├── pdf_generator.py         # Pembuat berkas PDF rekapitulasi formal (ReportLab) siap cetak & arsip
├── requirements.txt         # Daftar pustaka Python yang dibutuhkan
├── .gitignore               # Pengabaian berkas sementara & cache
├── .streamlit/
│   └── config.toml          # Konfigurasi tema dan server Streamlit
└── README.md                # Dokumentasi proyek
```

---

## 🛠️ Instalasi & Menjalankan Lokal

1. **Clone Repositori:**
   ```bash
   git clone https://github.com/darmayanimina-dev/FORMS-MINA.git
   cd FORMS-MINA
   ```

2. **Instal Dependensi:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Jalankan Aplikasi:**
   ```bash
   streamlit run app.py
   ```

Aplikasi dapat dibuka di browser melalui `http://localhost:8501` (atau port yang ditentukan).

---

## ☁️ Deployment ke Streamlit Community Cloud

1. Buka [share.streamlit.io](https://share.streamlit.io/).
2. Hubungkan akun GitHub Anda (`darmayanimina-dev`).
3. Pilih repositori: `darmayanimina-dev/FORMS-MINA`.
4. Branch: `main`.
5. Main file path: `app.py`.
6. Klik **Deploy**.
