import io
import re
import pandas as pd
import pdfplumber
import streamlit as st
from datetime import date

# --- 1. PENGATURAN TAMPILAN UTAMA ---
st.set_page_config(page_title="KODEX - Kompilator Dokumen", page_icon="🚢", layout="centered")

# CSS Kustom (Header diperkecil paddingnya)
st.markdown("""
    <style>
    .header-box { background: #f8fafc; padding: 0.8rem 1rem; border-radius: 8px; border-left: 4px solid #0284c7; margin-bottom: 0.5rem; }
    .stButton>button { width: 100%; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# --- 2. HEADER APLIKASI (Lebih Compact) ---
st.markdown("""
    <div class='header-box'>
        <h3 style='margin:0;'>🚢 KODEX (KOmpilator DOKumen EXim)</h3>
        <p style='margin:0; font-size: 0.9em; color:#475569;'>PT. Setia Samudera Abadi | Automasi CEISA 4.0</p>
    </div>
""", unsafe_allow_html=True)

# --- 3. INPUT PARAMETER (Tabs) ---
tab1, tab2, tab3 = st.tabs(["📌 Data Utama", "📥 Dokumen", "🚀 Proses"])

with tab1:
    jenis_pemberitahuan = st.radio("Pilih Jenis Dokumen:", ["PIB BC 2.0 (Impor)", "PEB BC 3.0 (Ekspor)"], horizontal=True)
    col1, col2 = st.columns(2)
    nomor_aju = col1.text_input("🔢 Nomor Aju", placeholder="000020PRO...")
    ndpbm_input = col2.number_input("💱 Nilai NDPBM", value=12644.5000, format="%.4f")

with tab2:
    c1, c2, c3 = st.columns(3)
    inv_file = c1.file_uploader("1. Invoice", type="pdf")
    pl_file = c2.file_uploader("2. Packing List", type="pdf")
    bl_file = c3.file_uploader("3. B/L / AWB", type="pdf")

with tab3:
    generate_btn = st.button("🚀 Generate Excel Sekarang", use_container_width=True)

# --- 4. LOGIKA EKSTRAKSI & PROSES ---
if generate_btn:
    if not inv_file or not pl_file or not nomor_aju:
        st.warning("⚠️ Mohon isi semua data di tab Data Utama dan unggah dokumen!")
    else:
        with st.spinner("Memproses data..."):
            try:
                # [EKSTRAKSI PDF]
                pl_text = "".join([p.extract_text() for p in pdfplumber.open(pl_file).pages])
                inv_text = "".join([p.extract_text() for p in pdfplumber.open(inv_file).pages])
                bl_text = "".join([p.extract_text() for p in pdfplumber.open(bl_file).pages]) if bl_file else ""
                
                val_netto = float(re.search(r"NET\s*WEIGHT\s*:?\s*([\d,.]+)", pl_text, re.IGNORECASE).group(1).replace(',', '')) if re.search(r"NET\s*WEIGHT\s*:?\s*([\d,.]+)", pl_text, re.IGNORECASE) else 0
                val_bruto = float(re.search(r"GROSS\s*WEIGHT\s*:?\s*([\d,.]+)", pl_text, re.IGNORECASE).group(1).replace(',', '')) if re.search(r"GROSS\s*WEIGHT\s*:?\s*([\d,.]+)", pl_text, re.IGNORECASE) else 0
                val_cif = float(re.search(r"(?:TOTAL|TOTAL VALUE)\s*:?\s*([\d,.]+)", inv_text, re.IGNORECASE).group(1).replace(',', '')) if re.search(r"(?:TOTAL|TOTAL VALUE)\s*:?\s*([\d,.]+)", inv_text, re.IGNORECASE) else 44443
                is_awb = "AWB" in bl_text.upper() or "AIR WAYBILL" in bl_text.upper()

                # --- SHEET HEADER ---
                header_cols = ['NOMOR AJU', 'KODE DOKUMEN', 'KODE KANTOR', 'KODE JENIS IMPOR', 'KODE JENIS PROSEDUR', 'KODE CARA BAYAR', 'TANGGAL PERNYATAAN', 'KOTA PERNYATAAN', 'KODE ASURANSI', 'KODE PELABUHAN TUJUAN', 'NAMA PERNYATAAN', 'JABATAN PERNYATAAN', 'FLAG PROPORSIONAL NETTO', 'ASURANSI', 'NILAI BARANG', 'BIAYA TAMBAHAN', 'BIAYA PENGURANG', 'VD', 'HARGA_PENYERAHAN', 'DASAR PENGENAAN PAJAK', 'UANG MUKA', 'VOLUME', 'PPN PAJAK', 'NETTO', 'BRUTO', 'NDPBM', 'FOB', 'CIF']
                df_header = pd.DataFrame(columns=header_cols)
                header_data = {'NOMOR AJU': nomor_aju, 'KODE DOKUMEN': 20, 'KODE KANTOR': "050100" if is_awb else "", 'KODE JENIS IMPOR': 1, 'KODE JENIS PROSEDUR': 1, 'KODE CARA BAYAR': 1, 'TANGGAL PERNYATAAN': date.today().strftime('%Y-%m-%d'), 'KOTA PERNYATAAN': 'BEKASI', 'KODE ASURANSI': 'DN', 'KODE PELABUHAN TUJUAN': 'IDCGK' if is_awb else "", 'NAMA PERNYATAAN': 'MUHAMMAD SUTAN ETHANOVA PRIMOLASSA', 'JABATAN PERNYATAAN': 'PELAKSANA', 'FLAG PROPORSIONAL NETTO': 'T', 'ASURANSI': 0, 'NILAI BARANG': 0, 'BIAYA TAMBAHAN': 0, 'BIAYA PENGURANG': 0, 'VD': 0, 'HARGA_PENYERAHAN': 0, 'DASAR PENGENAAN PAJAK': 0, 'UANG MUKA': 0, 'VOLUME': 0, 'PPN PAJAK': 0, 'NETTO': val_netto, 'BRUTO': val_bruto, 'NDPBM': ndpbm_input, 'FOB': 44443, 'CIF': val_cif}
                df_header = pd.concat([df_header, pd.DataFrame([header_data])], ignore_index=True)

                # --- INISIALISASI SHEET LAINNYA ---
                # (Disini Anda dapat meletakkan inisialisasi sheet lainnya yang ada di script lama Anda)
                # ... (seperti BAHANBAKUTARIF, PUNGUTAN, JAMINAN, dll) ...
                
                # --- EXCEL GENERATION ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    # Tulis semua sheet Anda di sini
                    df_header.to_excel(writer, sheet_name="HEADER", index=False)
                    # (Tambahkan df.to_excel untuk sheet lainnya...)

                st.success("✅ Excel berhasil di-generate!")
                st.download_button("⬇️ Download Excel", data=output.getvalue(), file_name=f"{nomor_aju}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"Error: {e}")

# Footer
st.markdown("<hr style='margin: 15px 0;'><div style='text-align: center; opacity: 0.5;'>© 2026 PT. Setia Samudera Abadi</div>", unsafe_allow_html=True)
