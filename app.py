import io
import re
import pandas as pd
import pdfplumber
import streamlit as st
from datetime import date

# --- 1. PENGATURAN TAMPILAN UTAMA ---
st.set_page_config(page_title="KODEX - Kompilator Dokumen", page_icon="🚢", layout="centered")

st.markdown("""<style>.block-container { padding-top: 1.5rem; max-width: 900px; } .stButton>button { width: 100%; }</style>""", unsafe_allow_html=True)

# --- 2. HEADER APLIKASI ---
st.markdown("## **KODEX** 🚢 <span style='font-size: 15px; opacity: 0.7;'>| PT. Setia Samudera Abadi</span>", unsafe_allow_html=True)
st.markdown("---")

# --- 3. INPUT PARAMETER ---
jenis_pemberitahuan = st.radio("📌 Pilih Jenis Dokumen:", ["PIB BC 2.0 (Impor)", "PEB BC 3.0 (Ekspor)"], horizontal=True)
col_in1, col_in2 = st.columns(2)
nomor_aju = col_in1.text_input("🔢 Masukkan Nomor Aju")
ndpbm_input = col_in2.number_input("💱 Masukkan Nilai NDPBM", value=12644.5000, format="%.4f")

# --- 4. UPLOAD DOKUMEN ---
c1, c2, c3 = st.columns(3)
inv_file = c1.file_uploader("1. Invoice (PDF)", type="pdf")
pl_file = c2.file_uploader("2. Packing List (PDF)", type="pdf")
bl_file = c3.file_uploader("3. House/Master B/L (PDF)", type="pdf")

generate_btn = st.button("🚀 Generate Excel CEISA 4.0", use_container_width=True)

# --- 5. LOGIKA EKSTRAKSI & PROSES ---
if generate_btn:
    if not inv_file or not pl_file or not nomor_aju:
        st.warning("⚠️ Mohon isi Nomor Aju dan unggah dokumen Invoice & Packing List.")
    else:
        with st.spinner("Mengolah data..."):
            try:
                # [EKSTRAKSI DATA PDF]
                pl_text = ""
                with pdfplumber.open(pl_file) as pdf:
                    for page in pdf.pages: pl_text += page.extract_text() + "\n"
                
                netto_match = re.search(r"NET\s*WEIGHT\s*:?\s*([\d,.]+)", pl_text, re.IGNORECASE)
                bruto_match = re.search(r"GROSS\s*WEIGHT\s*:?\s*([\d,.]+)", pl_text, re.IGNORECASE)
                val_netto = float(netto_match.group(1).replace(',', '')) if netto_match else 0
                val_bruto = float(bruto_match.group(1).replace(',', '')) if bruto_match else 0

                inv_text = ""
                with pdfplumber.open(inv_file) as pdf:
                    for page in pdf.pages: inv_text += page.extract_text() + "\n"
                cif_match = re.search(r"(?:TOTAL|TOTAL VALUE)\s*:?\s*([\d,.]+)", inv_text, re.IGNORECASE)
                val_cif = float(cif_match.group(1).replace(',', '')) if cif_match else 44443

                bl_text = ""
                if bl_file:
                    with pdfplumber.open(bl_file) as pdf:
                        bl_text = "".join([p.extract_text() for p in pdf.pages])
                is_awb = "AWB" in bl_text.upper() or "AIR WAYBILL" in bl_text.upper()

                # --- 6. INISIALISASI SHEET ---
                # A. HEADER
                header_cols = ['NOMOR AJU', 'KODE DOKUMEN', 'KODE KANTOR', 'KODE KANTOR BONGKAR', 'KODE KANTOR PERIKSA', 'KODE KANTOR TUJUAN', 'KODE KANTOR EKSPOR', 'KODE JENIS IMPOR', 'KODE JENIS EKSPOR', 'KODE JENIS TPB', 'KODE JENIS PLB', 'KODE JENIS PROSEDUR', 'KODE TUJUAN PEMASUKAN', 'KODE TUJUAN PENGIRIMAN', 'KODE TUJUAN TPB', 'KODE CARA DAGANG', 'KODE CARA BAYAR', 'KODE CARA BAYAR LAINNYA', 'KODE GUDANG ASAL', 'KODE GUDANG TUJUAN', 'KODE JENIS KIRIM', 'KODE JENIS PENGIRIMAN', 'KODE KATEGORI EKSPOR', 'KODE KATEGORI MASUK FTZ', 'KODE KATEGORI KELUAR FTZ', 'KODE KATEGORI BARANG FTZ', 'KODE LOKASI', 'KODE LOKASI BAYAR', 'LOKASI ASAL', 'LOKASI TUJUAN', 'KODE DAERAH ASAL', 'KODE NEGARA TUJUAN', 'KODE TUTUP PU', 'NOMOR BC11', 'TANGGAL BC11', 'NOMOR POS', 'NOMOR SUB POS', 'KODE PELABUHAN BONGKAR', 'KODE PELABUHAN MUAT', 'KODE PELABUHAN MUAT AKHIR', 'KODE PELABUHAN TRANSIT', 'KODE PELABUHAN TUJUAN', 'KODE PELABUHAN EKSPOR', 'KODE TPS', 'TANGGAL BERANGKAT', 'TANGGAL EKSPOR', 'TANGGAL MASUK', 'TANGGAL MUAT', 'TANGGAL TIBA', 'TANGGAL PERIKSA', 'TEMPAT STUFFING', 'TANGGAL STUFFING', 'KODE TANDA PENGAMAN', 'JUMLAH TANDA PENGAMAN', 'FLAG CURAH', 'FLAG SDA', 'FLAG VD', 'FLAG AP BK', 'FLAG MIGAS', 'KODE ASURANSI', 'ASURANSI', 'NILAI BARANG', 'NILAI INCOTERM', 'NILAI MAKLON', 'FREIGHT', 'FOB', 'BIAYA TAMBAHAN', 'BIAYA PENGURANG', 'VD', 'CIF', 'HARGA_PENYERAHAN', 'NDPBM', 'TOTAL DANA SAWIT', 'DASAR PENGENAAN PAJAK', 'NILAI JASA', 'UANG MUKA', 'BRUTO', 'NETTO', 'VOLUME', 'KOTA PERNYATAAN', 'TANGGAL PERNYATAAN', 'NAMA PERNYATAAN', 'JABATAN PERNYATAAN', 'KODE VALUTA', 'KODE INCOTERM', 'KODE JASA KENA PAJAK', 'NOMOR BUKTI BAYAR', 'TANGGAL BUKTI BAYAR', 'KODE JENIS NILAI', 'KODE KANTOR MUAT', 'NOMOR DAFTAR', 'TANGGAL DAFTAR', 'KODE ASAL BARANG FTZ', 'KODE TUJUAN PENGELUARAN', 'PPN PAJAK', 'PPNBM PAJAK', 'TARIF PPN PAJAK', 'TARIF PPNBM PAJAK', 'BARANG TIDAK BERWUJUD', 'KODE JENIS PENGELUARAN', 'BARANG KIRIMAN', 'FLAG KONSOL', 'KODE JENIS PENGANGKUTAN', 'FLAG PROPORSIONAL NETTO']
                df_header = pd.DataFrame(columns=header_cols)
                
                header_data = {
                    'NOMOR AJU': nomor_aju, 'KODE DOKUMEN': 20, 'KODE KANTOR': "050100" if is_awb else "",
                    'KODE JENIS IMPOR': 1, 'KODE JENIS PROSEDUR': 1, 'KODE CARA BAYAR': 1,
                    'TANGGAL PERNYATAAN': date.today().strftime('%Y-%m-%d'), 'KOTA PERNYATAAN': 'BEKASI',
                    'KODE ASURANSI': 'DN', 'KODE PELABUHAN TUJUAN': 'IDCGK' if is_awb else "",
                    'NAMA PERNYATAAN': 'MUHAMMAD SUTAN ETHANOVA PRIMOLASSA', 'JABATAN PERNYATAAN': 'PELAKSANA',
                    'FLAG PROPORSIONAL NETTO': 'T', 'ASURANSI': 0, 'NILAI BARANG': 0,
                    'BIAYA TAMBAHAN': 0, 'BIAYA PENGURANG': 0, 'VD': 0, 'HARGA_PENYERAHAN': 0,
                    'DASAR PENGENAAN PAJAK': 0, 'UANG MUKA': 0, 'VOLUME': 0, 'PPN PAJAK': 0,
                    'NETTO': val_netto, 'BRUTO': val_bruto, 'NDPBM': ndpbm_input, 
                    'FOB': 44443, 'CIF': val_cif
                }
                # PERBAIKAN: Menggunakan pd.concat untuk mengganti .append yang usang
                df_header = pd.concat([df_header, pd.DataFrame([header_data])], ignore_index=True)

                # B. SHEET LAINNYA
                df_bahanbakutarif = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI BAHAN BAKU', 'KODE ASAL BAHAN BAKU', 'KODE PUNGUTAN', 'KODE TARIF', 'TARIF', 'KODE FASILITAS', 'TARIF FASILITAS', 'NILAI BAYAR', 'NILAI FASILITAS', 'NILAI SUDAH DILUNASI', 'KODE SATUAN', 'JUMLAH SATUAN', 'FLAG BMT SEMENTARA', 'KODE KOMODITI CUKAI', 'KODE SUB KOMODITI CUKAI', 'FLAG TIS', 'FLAG PELEKATAN', 'KODE KEMASAN', 'KODE SUB KOMODITI CUKAI', 'FLAG TIS', 'FLAG PELEKATAN', 'KODE KEMASAN', 'JUMLAH KEMASAN'])
                df_bahanbakudokumen = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI BAHAN BAKU', 'KODE_ASAL_BAHAN_BAKU', 'SERI DOKUMEN', 'SERI IZIN'])
                df_pungutan = pd.DataFrame(columns=['NOMOR AJU', 'KODE FASILITAS TARIF', 'KODE JENIS PUNGUTAN', 'NILAI PUNGUTAN', 'NPWP BILLING'])
                df_jaminan = pd.DataFrame(columns=['NOMOR AJU', 'KODE KANTOR', 'KODE JAMINAN', 'NOMOR JAMINAN', 'TANGGAL JAMINAN', 'NILAI JAMINAN', 'PENJAMIN', 'TANGGAL JATUH TEMPO', 'NOMOR BPJ', 'TANGGAL BPJ'])
                
                # Sheet lain (default kosong)
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

                # --- 7. EXCEL GENERATION ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    sheets = [
                        ("HEADER", df_header), ("ENTITAS", df_entitas), ("DOKUMEN", df_dokumen),
                        ("PENGANGKUT", df_pengangkut), ("KEMASAN", df_kemasan), ("KONTAINER", df_kontainer),
                        ("KOMPONENBIAYA", df_komponenbiaya), ("BARANG", df_barang), ("BARANGTARIF", df_barangtarif),
                        ("BARANGDOKUMEN", df_barangdokumen), ("BARANGENTITAS", df_barangentitas),
                        ("BARANGSPEKKHUSUS", df_barangspekkhusus), ("BARANGVD", df_barangvd), ("BAHANBAKU", df_bahanbaku),
                        ("BAHANBAKUTARIF", df_bahanbakutarif), ("BAHANBAKUDOKUMEN", df_bahanbakudokumen),
                        ("PUNGUTAN", df_pungutan), ("JAMINAN", df_jaminan), ("BANKDEVISA", df_bankdevisa),
                        ("VERSI", df_versi), ("RESPON", df_respon)
                    ]
                    for name, df in sheets: df.to_excel(writer, sheet_name=name, index=False)

                st.success("✅ Excel berhasil di-generate!")
                st.download_button("⬇️ Download Excel", data=output.getvalue(), file_name=f"{nomor_aju}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"Terjadi kesalahan teknis: {e}")

# Footer
st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
st.markdown("<div style='text-align: center; font-size: 11px; opacity: 0.6;'>© 2026 PT. Setia Samudera Abadi &bull; v1.0.0</div>", unsafe_allow_html=True)
