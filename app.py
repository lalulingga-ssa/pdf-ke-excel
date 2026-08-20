import io
import re
import pandas as pd
import pdfplumber
import streamlit as st
import os
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

# --- 3. DATABASE HELPER ---
DB_FILE = "database_customer.csv"

def load_db():
    if os.path.exists(DB_FILE):
        return pd.read_csv(DB_FILE)
    return pd.DataFrame(columns=["Tipe", "Nama", "Alamat", "No Identitas", "NIB Entitas", "Kode Negara"])

def save_db(df):
    df.to_csv(DB_FILE, index=False)

# --- 4. TABS & LAYOUT ---
tab1, tab2, tab3, tab4 = st.tabs(["📌 Data Utama", "📥 Dokumen", "🗄️ Database Customer", "🚀 Proses"])

with tab1:
    jenis_pemberitahuan = st.radio("Pilih Jenis Dokumen:", ["PIB BC 2.0 (Impor)", "PEB BC 3.0 (Ekspor)"], horizontal=True)
    col1, col2 = st.columns(2)
    nomor_aju = col1.text_input("🔢 Nomor Aju", placeholder="Contoh: 000020PRO325...")
    ndpbm_input = col2.number_input("💱 Nilai NDPBM", value=12644.5000, format="%.4f")

with tab2:
    inv_file = st.file_uploader("1. Invoice", type="pdf")
    pl_file = st.file_uploader("2. Packing List", type="pdf")
    hbl_file = st.file_uploader("3. House B/L", type="pdf")
    mbl_file = st.file_uploader("4. Master B/L", type="pdf")
    bc_file = st.file_uploader("5. Manifest BC 1.1", type="pdf")

with tab3:
    st.subheader("Tambah Importir/Shipper Baru")
    with st.form("customer_form"):
        tipe_cust = st.selectbox("Tipe:", ["Shipper", "Consignee"])
        nama_cust = st.text_input("Nama (Importir/Shipper):")
        alamat_cust = st.text_area("Alamat:")
        id_cust = st.text_input("No Identitas:")
        nib_cust = st.text_input("NIB Entitas:")
        negara_cust = st.text_input("Kode Negara (Contoh: ID, CN, AU):")
        if st.form_submit_button("💾 Simpan Customer"):
            new_data = pd.DataFrame([[tipe_cust, nama_cust, alamat_cust, id_cust, nib_cust, negara_cust]], 
                                    columns=["Tipe", "Nama", "Alamat", "No Identitas", "NIB Entitas", "Kode Negara"])
            save_db(pd.concat([load_db(), new_data], ignore_index=True))
            st.success("Customer disimpan!")
    
    st.divider()
    st.subheader("Database Tersimpan")
    st.dataframe(load_db(), use_container_width=True)

with tab4:
    generate_btn = st.button("🚀 Generate Excel Sekarang", use_container_width=True)

# --- 5. LOGIKA EKSTRAKSI & PROSES ---
if generate_btn:
    if not inv_file or not pl_file or not nomor_aju:
        st.warning("⚠️ Mohon isi semua data di tab Data Utama dan unggah dokumen!")
    else:
        with st.spinner("Memproses data..."):
            try:
                # [EKSTRAKSI DATA PDF]
                pl_text = "".join([p.extract_text() for p in pdfplumber.open(pl_file).pages]) if pl_file else ""
                inv_text = "".join([p.extract_text() for p in pdfplumber.open(inv_file).pages]) if inv_file else ""
                bl_text = "".join([p.extract_text() for p in pdfplumber.open(hbl_file).pages]) if hbl_file else ""
                
                val_netto = float(re.search(r"NET\s*WEIGHT\s*:?\s*([\d,.]+)", pl_text, re.IGNORECASE).group(1).replace(',', '')) if re.search(r"NET\s*WEIGHT\s*:?\s*([\d,.]+)", pl_text, re.IGNORECASE) else 0
                val_bruto = float(re.search(r"GROSS\s*WEIGHT\s*:?\s*([\d,.]+)", pl_text, re.IGNORECASE).group(1).replace(',', '')) if re.search(r"GROSS\s*WEIGHT\s*:?\s*([\d,.]+)", pl_text, re.IGNORECASE) else 0
                val_cif = float(re.search(r"(?:TOTAL|TOTAL VALUE)\s*:?\s*([\d,.]+)", inv_text, re.IGNORECASE).group(1).replace(',', '')) if re.search(r"(?:TOTAL|TOTAL VALUE)\s*:?\s*([\d,.]+)", inv_text, re.IGNORECASE) else 44443
                is_awb = "AWB" in bl_text.upper() or "AIR WAYBILL" in bl_text.upper()

                # --- SHEET HEADER ---
                header_data = {'NOMOR AJU': nomor_aju, 'KODE DOKUMEN': 20, 'KODE KANTOR': "050100" if is_awb else "", 'TANGGAL PERNYATAAN': date.today().strftime('%Y-%m-%d'), 'KOTA PERNYATAAN': 'BEKASI', 'NAMA PERNYATAAN': 'MUHAMMAD SUTAN ETHANOVA PRIMOLASSA', 'JABATAN PERNYATAAN': 'PELAKSANA', 'NETTO': val_netto, 'BRUTO': val_bruto, 'NDPBM': ndpbm_input, 'FOB': 44443, 'CIF': val_cif}
                df_header = pd.concat([pd.DataFrame(columns=['NOMOR AJU', 'KODE DOKUMEN', 'KODE KANTOR', 'TANGGAL PERNYATAAN', 'KOTA PERNYATAAN', 'NAMA PERNYATAAN', 'JABATAN PERNYATAAN', 'NETTO', 'BRUTO', 'NDPBM', 'FOB', 'CIF']), pd.DataFrame([header_data])], ignore_index=True)

                # --- SHEET ENTITAS ---
                # Struktur Entitas sudah termasuk KODE NEGARA
                df_entitas = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'KODE ENTITAS', 'KODE JENIS IDENTITAS', 'KODE NEGARA'])
                data_entitas = {
                    'NOMOR AJU': [nomor_aju]*7,
                    'SERI': [10, 9, 11, 1, 7, 4, 10],
                    'KODE ENTITAS': [10, 9, 11, 1, 7, 4, 10],
                    'KODE JENIS IDENTITAS': [None, None, None, 6, 6, 6, 6],
                    'KODE NEGARA': ['ID']*7 # Default ID, bisa disesuaikan dari DB
                }
                df_entitas = pd.concat([df_entitas, pd.DataFrame(data_entitas)], ignore_index=True)

                # --- INISIALISASI SHEET LAINNYA ---
                df_bahanbakutarif = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI BAHAN BAKU', 'KODE ASAL BAHAN BAKU', 'KODE PUNGUTAN', 'KODE TARIF', 'TARIF', 'KODE FASILITAS', 'TARIF FASILITAS', 'NILAI BAYAR', 'NILAI FASILITAS', 'NILAI SUDAH DILUNASI', 'KODE SATUAN', 'JUMLAH SATUAN', 'FLAG BMT SEMENTARA', 'KODE KOMODITI CUKAI', 'KODE SUB KOMODITI CUKAI', 'FLAG TIS', 'FLAG PELEKATAN', 'KODE KEMASAN', 'KODE SUB KOMODITI CUKAI', 'FLAG TIS', 'FLAG PELEKATAN', 'KODE KEMASAN', 'JUMLAH KEMASAN'])
                df_bahanbakudokumen = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI BAHAN BAKU', 'KODE_ASAL_BAHAN_BAKU', 'SERI DOKUMEN', 'SERI IZIN'])
                df_pungutan = pd.DataFrame(columns=['NOMOR AJU', 'KODE FASILITAS TARIF', 'KODE JENIS PUNGUTAN', 'NILAI PUNGUTAN', 'NPWP BILLING'])
                df_jaminan = pd.DataFrame(columns=['NOMOR AJU', 'KODE KANTOR', 'KODE JAMINAN', 'NOMOR JAMINAN', 'TANGGAL JAMINAN', 'NILAI JAMINAN', 'PENJAMIN', 'TANGGAL JATUH TEMPO', 'NOMOR BPJ', 'TANGGAL BPJ'])
                
                # --- EXCEL GENERATION ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    sheets = [("HEADER", df_header), ("ENTITAS", df_entitas), ("BAHANBAKUTARIF", df_bahanbakutarif), ("BAHANBAKUDOKUMEN", df_bahanbakudokumen), ("PUNGUTAN", df_pungutan), ("JAMINAN", df_jaminan)]
                    for name, df in sheets: df.to_excel(writer, sheet_name=name, index=False)

                st.success("✅ Excel berhasil di-generate!")
                st.download_button("⬇️ Download Excel", data=output.getvalue(), file_name=f"{nomor_aju}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"Error: {e}")

# Footer
st.markdown("<hr style='margin: 15px 0;'><div style='text-align: center; opacity: 0.5;'>© 2026 PT. Setia Samudera Abadi</div>", unsafe_allow_html=True)
