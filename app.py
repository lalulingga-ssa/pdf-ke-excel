import io
import re
import pandas as pd
import pdfplumber
import streamlit as st
from datetime import date

# --- 1. PENGATURAN TAMPILAN ---
st.set_page_config(page_title="KODEX - Kompilator Dokumen", page_icon="🚢", layout="centered")

st.markdown("""
    <style>
    .header-box { background: #f8fafc; padding: 1rem; border-radius: 8px; border-left: 5px solid #0284c7; margin-bottom: 1rem; color: #1e293b; }
    .stButton>button { width: 100%; font-weight: bold; background: #0284c7; color: white; }
    </style>
""", unsafe_allow_html=True)

# --- 2. HEADER ---
st.markdown("""
    <div class='header-box'>
        <h3 style='margin:0; color:#1e293b;'>🚢 KODEX (KOmpilator DOKumen EXim)</h3>
        <p style='margin:0; font-size: 0.9em; color:#475569;'>PT. Setia Samudera Abadi | Automasi CEISA 4.0</p>
    </div>
""", unsafe_allow_html=True)

# --- 3. INPUT (TABS) ---
tab1, tab2, tab3 = st.tabs(["📌 Data Utama", "📥 Dokumen", "🚀 Proses"])

with tab1:
    jenis_pemberitahuan = st.radio("Pilih Jenis Dokumen:", ["PIB BC 2.0 (Impor)", "PEB BC 3.0 (Ekspor)"], horizontal=True)
    col1, col2 = st.columns(2)
    nomor_aju = col1.text_input("🔢 Nomor Aju", placeholder="Contoh: 000020PRO325...")
    ndpbm_input = col2.number_input("💱 Nilai NDPBM", value=12644.5000, format="%.4f")

with tab2:
    # Memisahkan kembali kolom upload agar tidak digabung
    c1, c2, c3 = st.columns(3)
    inv_file = c1.file_uploader("1. Invoice", type="pdf")
    pl_file = c2.file_uploader("2. Packing List", type="pdf")
    hbl_file = c3.file_uploader("3. House B/L", type="pdf")
    
    c4, c5 = st.columns(2)
    mbl_file = c4.file_uploader("4. Master B/L", type="pdf")
    bc_file = c5.file_uploader("5. Manifest BC 1.1", type="pdf")

with tab3:
    generate_btn = st.button("🚀 Generate Excel Sekarang", use_container_width=True)

# --- 4. LOGIKA EKSTRAKSI & PROSES ---
if generate_btn:
    if not inv_file or not pl_file or not nomor_aju:
        st.warning("⚠️ Mohon isi Nomor Aju dan unggah Invoice & Packing List!")
    else:
        with st.spinner("Memproses data..."):
            try:
                # [EKSTRAKSI DATA]
                pl_text = "".join([p.extract_text() for p in pdfplumber.open(pl_file).pages]) if pl_file else ""
                inv_text = "".join([p.extract_text() for p in pdfplumber.open(inv_file).pages]) if inv_file else ""
                bl_text = "".join([p.extract_text() for p in pdfplumber.open(hbl_file).pages]) if hbl_file else ""
                
                val_netto = float(re.search(r"NET\s*WEIGHT\s*:?\s*([\d,.]+)", pl_text, re.IGNORECASE).group(1).replace(',', '')) if re.search(r"NET\s*WEIGHT\s*:?\s*([\d,.]+)", pl_text, re.IGNORECASE) else 0
                val_bruto = float(re.search(r"GROSS\s*WEIGHT\s*:?\s*([\d,.]+)", pl_text, re.IGNORECASE).group(1).replace(',', '')) if re.search(r"GROSS\s*WEIGHT\s*:?\s*([\d,.]+)", pl_text, re.IGNORECASE) else 0
                val_cif = float(re.search(r"(?:TOTAL|TOTAL VALUE)\s*:?\s*([\d,.]+)", inv_text, re.IGNORECASE).group(1).replace(',', '')) if re.search(r"(?:TOTAL|TOTAL VALUE)\s*:?\s*([\d,.]+)", inv_text, re.IGNORECASE) else 44443
                is_awb = "AWB" in bl_text.upper() or "AIR WAYBILL" in bl_text.upper()

                # --- SHEET HEADER ---
                header_data = {
                    'NOMOR AJU': nomor_aju, 'KODE DOKUMEN': 20, 'KODE KANTOR': "050100" if is_awb else "",
                    'TANGGAL PERNYATAAN': date.today().strftime('%Y-%m-%d'), 'KOTA PERNYATAAN': 'BEKASI',
                    'NAMA PERNYATAAN': 'MUHAMMAD SUTAN ETHANOVA PRIMOLASSA', 'JABATAN PERNYATAAN': 'PELAKSANA',
                    'NETTO': val_netto, 'BRUTO': val_bruto, 'NDPBM': ndpbm_input, 'FOB': 44443, 'CIF': val_cif
                }
                cols = ['NOMOR AJU', 'KODE DOKUMEN', 'KODE KANTOR', 'KODE JENIS IMPOR', 'KODE JENIS PROSEDUR', 'KODE CARA BAYAR', 'TANGGAL PERNYATAAN', 'KOTA PERNYATAAN', 'KODE ASURANSI', 'KODE PELABUHAN TUJUAN', 'NAMA PERNYATAAN', 'JABATAN PERNYATAAN', 'FLAG PROPORSIONAL NETTO', 'ASURANSI', 'NILAI BARANG', 'BIAYA TAMBAHAN', 'BIAYA PENGURANG', 'VD', 'HARGA_PENYERAHAN', 'DASAR PENGENAAN PAJAK', 'UANG MUKA', 'VOLUME', 'PPN PAJAK', 'NETTO', 'BRUTO', 'NDPBM', 'FOB', 'CIF']
                df_header = pd.concat([pd.DataFrame(columns=cols), pd.DataFrame([header_data])], ignore_index=True)

                # --- INISIALISASI SHEET LAINNYA ---
                df_bahanbakutarif = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI BAHAN BAKU', 'KODE ASAL BAHAN BAKU', 'KODE PUNGUTAN', 'KODE TARIF', 'TARIF', 'KODE FASILITAS', 'TARIF FASILITAS', 'NILAI BAYAR', 'NILAI FASILITAS', 'NILAI SUDAH DILUNASI', 'KODE SATUAN', 'JUMLAH SATUAN', 'FLAG BMT SEMENTARA', 'KODE KOMODITI CUKAI', 'KODE SUB KOMODITI CUKAI', 'FLAG TIS', 'FLAG PELEKATAN', 'KODE KEMASAN', 'KODE SUB KOMODITI CUKAI', 'FLAG TIS', 'FLAG PELEKATAN', 'KODE KEMASAN', 'JUMLAH KEMASAN'])
                df_bahanbakudokumen = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI BAHAN BAKU', 'KODE_ASAL_BAHAN_BAKU', 'SERI DOKUMEN', 'SERI IZIN'])
                df_pungutan = pd.DataFrame(columns=['NOMOR AJU', 'KODE FASILITAS TARIF', 'KODE JENIS PUNGUTAN', 'NILAI PUNGUTAN', 'NPWP BILLING'])
                df_jaminan = pd.DataFrame(columns=['NOMOR AJU', 'KODE KANTOR', 'KODE JAMINAN', 'NOMOR JAMINAN', 'TANGGAL JAMINAN', 'NILAI JAMINAN', 'PENJAMIN', 'TANGGAL JATUH TEMPO', 'NOMOR BPJ', 'TANGGAL BPJ'])
                df_entitas = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'KODE ENTITAS', 'KODE JENIS IDENTITAS', 'NOMOR IDENTITAS', 'NAMA ENTITAS', 'ALAMAT ENTITAS', 'NIB ENTITAS', 'KODE JENIS API', 'KODE STATUS', 'NOMOR IJIN ENTITAS', 'TANGGAL IJIN ENTITAS', 'KODE NEGARA', 'NIPER ENTITAS', 'KODE KATEGORI KONSOLIDATOR', 'KODE AFILIASI'])
                df_dokumen = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'KODE DOKUMEN', 'NOMOR DOKUMEN', 'TANGGAL DOKUMEN', 'KODE FASILITAS', 'KODE IJIN'])
                df_pengangkut = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'KODE CARA ANGKUT', 'NAMA PENGANGKUT', 'NOMOR PENGANGKUT', 'KODE BENDERA', 'CALL SIGN', 'FLAG ANGKUT PLB', 'CARA PENGANGKUTAN LAINNYA'])
                df_kemasan = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'KODE KEMASAN', 'JUMLAH KEMASAN', 'MEREK', 'NOMOR SEGEL'])
                df_kontainer = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'NOMOR KONTINER', 'KODE UKURAN KONTAINER', 'KODE JENIS KONTAINER', 'KODE TIPE KONTAINER', 'NOMOR SEGEL'])
                df_komponenbiaya = pd.DataFrame(columns=['NOMOR AJU', 'JENIS NILAI', 'HARGA INVOICE', 'PEMBAYARAN TIDAK LANGSUNG', 'DISKON', 'KOMISI PENJUALAN', 'BIAYA PENGEMASAN', 'BIAYA PENGEPAKAN', 'ASSIST', 'ROYALTI', 'PROCEEDS', 'BIAYA TRANSPORTASI', 'BIAYA PEMUATAN', 'ASURANSI', 'GARANSI', 'BIAYA KEPENTINGAN SENDIRI', 'BIAYA PASCA IMPOR', 'BIAYA PAJAK INTERNAL', 'BUNGA', 'DEVIDEN'])
                df_barang = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'HS', 'KODE BARANG', 'URAIAN', 'MEREK', 'TIPE', 'UKURAN', 'SPESIFIKASI LAIN', 'KODE SATUAN', 'JUMLAH SATUAN', 'KODE KEMASAN', 'JUMLAH KEMASAN', 'KODE DOKUMEN ASAL', 'KODE KANTOR ASAL', 'NOMOR DAFTAR ASAL', 'TANGGAL DAFTAR ASAL', 'NOMOR AJU ASAL', 'SERI BARANG ASAL', 'NETTO', 'BRUTO', 'VOLUME', 'SALDO AWAL', 'SALDO AKHIR', 'JUMLAH REALISASI', 'CIF', 'CIF RUPIAH', 'NDPBM', 'FOB', 'ASURANSI', 'FREIGHT', 'NILAI TAMBAH', 'DISKON', 'HARGA PENYERAHAN', 'HARGA PEROLEHAN', 'HARGA SATUAN', 'HARGA EKSPOR', 'HARGA PATOKAN', 'NILAI BARANG', 'NILAI JASA', 'NILAI DANA SAWIT', 'NILAI DEVISA', 'PERSENTASE IMPOR', 'KODE ASAL BARANG', 'KODE DAERAH ASAL', 'KODE GUNA BARANG', 'KODE JENIS NILAI', 'JATUH TEMPO ROYALTI', 'KODE KATEGORI BARANG', 'KODE KONDISI BARANG', 'KODE NEGARA ASAL', 'KODE PERHITUNGAN', 'PERNYATAAN LARTAS', 'FLAG 4 TAHUN', 'SERI IZIN', 'TAHUN PEMBUATAN', 'KAPASITAS SILINDER', 'KODE BKC', 'KODE KOMODITI BKC', 'KODE SUB KOMODITI BKC', 'FLAG TIS', 'ISI PER KEMASAN', 'JUMLAH DILEKATKAN', 'JUMLAH PITA CUKAI', 'HJE CUKAI', 'TARIF CUKAI', 'KODE JENIS EKSPOR', 'METODE PENENTUAN NILAI', 'ALASAN METODE PENENTUAN NILAI', 'STATEMENT PERBEDAAN HARGA'])
                df_barangtarif = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'KODE PUNGUTAN', 'KODE TARIF', 'TARIF', 'KODE FASILITAS', 'TARIF FASILITAS', 'NILAI BAYAR', 'NILAI FASILITAS', 'NILAI SUDAH DILUNASI', 'KODE SATUAN', 'JUMLAH SATUAN', 'FLAG BMT SEMENTARA', 'KODE KOMODITI CUKAI', 'KODE SUB KOMODITI CUKAI', 'FLAG TIS', 'FLAG PELEKATAN', 'KODE KEMASAN', 'JUMLAH KEMASAN'])
                df_barangdokumen = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI DOKUMEN', 'SERI IZIN'])
                df_barangentitas = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI ENTITAS'])
                df_barangspekkhusus = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'KODE', 'URAIAN'])
                df_barangvd = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'KODE VD', 'NILAI BARANG', 'BIAYA TAMBAHAN', 'BIAYA PENGURANG', 'JATUH TEMPO'])
                df_bahanbaku = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI BAHAN BAKU', 'KODE ASAL BAHAN BAKU', 'HS', 'KODE BARANG', 'URAIAN', 'MEREK', 'TIPE', 'UKURAN', 'SPESIFIKASI LAIN', 'KODE SATUAN', 'JUMLAH SATUAN', 'KODE KEMASAN', 'JUMLAH KEMASAN', 'KODE DOKUMEN ASAL', 'KODE KANTOR ASAL', 'NOMOR DAFTAR ASAL', 'TANGGAL DAFTAR ASAL', 'NOMOR AJU ASAL', 'SERI BARANG ASAL', 'NETTO', 'BRUTO', 'VOLUME', 'CIF', 'CIF RUPIAH', 'NDPBM', 'HARGA PENYERAHAN', 'HARGA PEROLEHAN', 'NILAI JASA', 'SERI IZIN', 'VALUTA', 'KODE BKC', 'KODE KOMODITI BKC', 'KODE SUB KOMODITI BKC', 'FLAG TIS', 'ISI PER KEMASAN', 'JUMLAH DILEKATKAN', 'JUMLAH PITA CUKAI', 'HJE CUKAI', 'TARIF CUKAI'])
                df_bankdevisa = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'KODE', 'NAMA'])
                df_versi = pd.DataFrame({'VERSI': [1.3]})
                df_respon = pd.DataFrame(columns=['NOMOR AJU', 'KODE RESPON', 'NOMOR RESPON', 'TANGGAL RESPON'])

                # --- EXCEL GENERATION ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    sheets = [("HEADER", df_header), ("ENTITAS", df_entitas), ("DOKUMEN", df_dokumen), ("PENGANGKUT", df_pengangkut), ("KEMASAN", df_kemasan), ("KONTAINER", df_kontainer), ("KOMPONENBIAYA", df_komponenbiaya), ("BARANG", df_barang), ("BARANGTARIF", df_barangtarif), ("BARANGDOKUMEN", df_barangdokumen), ("BARANGENTITAS", df_barangentitas), ("BARANGSPEKKHUSUS", df_barangspekkhusus), ("BARANGVD", df_barangvd), ("BAHANBAKU", df_bahanbaku), ("BAHANBAKUTARIF", df_bahanbakutarif), ("BAHANBAKUDOKUMEN", df_bahanbakudokumen), ("PUNGUTAN", df_pungutan), ("JAMINAN", df_jaminan), ("BANKDEVISA", df_bankdevisa), ("VERSI", df_versi), ("RESPON", df_respon)]
                    for name, df in sheets: df.to_excel(writer, sheet_name=name, index=False)
                
                st.success("✅ Excel berhasil di-generate!")
                st.download_button("⬇️ Download Excel", data=output.getvalue(), file_name=f"{nomor_aju}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"Error: {e}")

# Footer
st.markdown("<hr style='margin: 15px 0;'><div style='text-align: center; opacity: 0.5;'>© 2026 PT. Setia Samudera Abadi</div>", unsafe_allow_html=True)
