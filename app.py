import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# --- 1. PENGATURAN TAMPILAN UTAMA ---
st.set_page_config(page_title="KODEX - PT. Setia Samudera Abadi", page_icon="🚢", layout="wide")

# CSS Kustom untuk tampilan yang elegan, bersih, dan profesional
st.markdown("""
    <style>
    .main { background-color: #f4f6f9; }
    h1 { color: #0a3d62; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-weight: 800; }
    h3 { color: #3c6382; }
    .stButton>button {
        background-color: #0a3d62; color: white; border-radius: 6px;
        padding: 0.6rem 1rem; border: none; transition: 0.3s; font-weight: bold;
    }
    .stButton>button:hover { background-color: #38ada9; color: white; border: none; }
    </style>
""", unsafe_allow_html=True)

# --- 2. HEADER APLIKASI ---
st.title("KODEX 🚢")
st.markdown("### *Kompilator Dokumen Ekspor-Impor | PT. Setia Samudera Abadi*")
st.write("Unggah dokumen kepabeanan Anda. Sistem akan mengekstrak informasi untuk disatukan menjadi satu file Excel template CEISA 4.0.")
st.markdown("---")

# --- 3. INPUT NILAI KURS ---
st.subheader("💱 Pengaturan Nilai Kurs")
ndpbm_rate = st.number_input("Masukkan Nilai NDPBM (Kurs Pajak Hari Ini dalam Rupiah):", min_value=1.0, value=15000.0, step=100.0)
st.markdown("---")

# --- 4. AREA UPLOAD ---
st.subheader("📥 Unggah Dokumen Pendukung")

col1, col2, col3 = st.columns(3)
with col1:
    inv_file = st.file_uploader("1. Invoice", type="pdf")
with col2:
    pl_file = st.file_uploader("2. Packing List", type="pdf")
with col3:
    hbl_file = st.file_uploader("3. House B/L", type="pdf")

col4, col5, col6 = st.columns(3)
with col4:
    mbl_file = st.file_uploader("4. Master B/L", type="pdf")
with col5:
    bc_file = st.file_uploader("5. Manifest BC 1.1", type="pdf")
with col6:
    st.info("💡 Pastikan format tabel di dalam dokumen adalah versi final dan tidak dienkripsi.")

st.markdown("---")

# --- 5. DAFTAR 70 KOLOM CEISA ---
ceisa_columns = [
    'NOMOR AJU', 'SERI BARANG', 'HS', 'KODE BARANG', 'URAIAN', 'MEREK', 'TIPE', 'UKURAN', 'SPESIFIKASI LAIN', 
    'KODE SATUAN', 'JUMLAH SATUAN', 'KODE KEMASAN', 'JUMLAH KEMASAN', 'KODE DOKUMEN ASAL', 'KODE KANTOR ASAL', 
    'NOMOR DAFTAR ASAL', 'TANGGAL DAFTAR ASAL', 'NOMOR AJU ASAL', 'SERI BARANG ASAL', 'NETTO', 'BRUTO', 'VOLUME', 
    'SALDO AWAL', 'SALDO AKHIR', 'JUMLAH REALISASI', 'CIF', 'CIF RUPIAH', 'NDPBM', 'FOB', 'ASURANSI', 'FREIGHT', 
    'NILAI TAMBAH', 'DISKON', 'HARGA PENYERAHAN', 'HARGA PEROLEHAN', 'HARGA SATUAN', 'HARGA EKSPOR', 'HARGA PATOKAN', 
    'NILAI BARANG', 'NILAI JASA', 'NILAI DANA SAWIT', 'NILAI DEVISA', 'PERSENTASE IMPOR', 'KODE ASAL BARANG', 
    'KODE DAERAH ASAL', 'KODE GUNA BARANG', 'KODE JENIS NILAI', 'JATUH TEMPO ROYALTI', 'KODE KATEGORI BARANG', 
    'KODE KONDISI BARANG', 'KODE NEGARA ASAL', 'KODE PERHITUNGAN', 'PERNYATAAN LARTAS', 'FLAG 4 TAHUN', 'SERI IZIN', 
    'TAHUN PEMBUATAN', 'KBerdasarkan gambar **Screenshot 2026-08-19 at 10.50.12.jpg** yang Anda lampirkan, masalahnya bukan karena script lama masih terbawa, melainkan ada bagian kode yang terpotong saat Anda menimpa/melakukan *copy-paste* script baru tersebut.

Pesan error yang muncul adalah:
> `SyntaxError: '[' was never closed` di file `app.py` baris 57 pada bagian `ceisa_columns = [`

### Penyebab Error
Di dalam bahasa pemrograman Python, tanda kurung siku buka `[` digunakan untuk membuat sebuah daftar (*list*). Error ini terjadi karena Anda membuka *list* dengan `[`, tetapi sistem tidak menemukan tanda kurung siku tutup `]` pasangannya di akhir daftar tersebut.

### Cara Memperbaikinya
Silakan buka kembali file `app.py` Anda, lalu gulir ke **baris 57**. Periksa deklarasi variabel `ceisa_columns`. Pastikan pada bagian akhir daftar kolom tersebut, Anda menutupnya dengan tanda `]`.

Bentuknya seharusnya seperti ini:

```python
ceisa_columns = [
    "Nama Kolom 1",
    "Nama Kolom 2",
    "Nama Kolom 3"
]  # <--- Pastikan tanda kurung siku tutup ini ada di akhir list
