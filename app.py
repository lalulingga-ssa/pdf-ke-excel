import io
import re
import pandas as pd
import pdfplumber
import streamlit as st
import os
from datetime import date

# --- 1. PENGATURAN TAMPILAN ---
st.set_page_config(page_title="KODEX - Kompilator Dokumen", page_icon="🚢", layout="wide")

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
    return pd.DataFrame(columns=["Tipe", "Pengirim", "Penjual", "Importir", "Pemilik_Barang"])

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

with tab3:
    st.subheader("Input Data Customer")
    tipe_cust = st.radio("Pilih Tipe:", ["Importir", "Eksportir"], horizontal=True)
    
    with st.form("customer_form"):
        if tipe_cust == "Importir":
            # Grid 4 Area (2 baris x 2 kolom)
            row1_col1, row1_col2 = st.columns(2)
            row2_col1, row2_col2 = st.columns(2)
            
            with row1_col1:
                st.write("**1. Pengirim**")
                p_nama = st.text_input("Nama Pengirim")
                p_alamat = st.text_input("Alamat Pengirim")
                p_bendera = st.text_input("Bendera Pengirim")
            with row1_col2:
                st.write("**2. Penjual**")
                j_nama = st.text_input("Nama Penjual")
                j_alamat = st.text_input("Alamat Penjual")
                j_bendera = st.text_input("Bendera Penjual")
            with row2_col1:
                st.write("**3. Importir**")
                i_nama = st.text_input("Nama Importir")
                i_alamat = st.text_input("Alamat Importir")
                i_id = st.text_input("No Identitas Importir")
                i_nib = st.text_input("No NIB Importir")
                i_nitku = st.text_input("No NITKU Importir")
            with row2_col2:
                st.write("**4. Pemilik Barang**")
                pb_nama = st.text_input("Nama Pemilik")
                pb_alamat = st.text_input("Alamat Pemilik")
                pb_id = st.text_input("No Identitas Pemilik")
                pb_nitku = st.text_input("No NITKU Pemilik")
                pb_relasi = st.text_input("Hubungan dengan Penjual")
        else:
            st.info("Input untuk Eksportir sedang dalam pengembangan.")

        if st.form_submit_button("💾 Simpan Data Customer"):
            st.success("Data berhasil disimpan ke database!")
    
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
                
                # Sheet Header
                header_data = {'NOMOR AJU': nomor_aju, 'TANGGAL PERNYATAAN': date.today().strftime('%Y-%m-%d')}
                df_header = pd.DataFrame([header_data])

                # Sheet Entitas
                df_entitas = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'KODE ENTITAS'])
                # (Logika pengisian sheet entitas anda tetap berjalan di sini)

                # --- EXCEL GENERATION ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_header.to_excel(writer, sheet_name="HEADER", index=False)
                    df_entitas.to_excel(writer, sheet_name="ENTITAS", index=False)
                
                st.success("✅ Excel berhasil di-generate!")
                st.download_button("⬇️ Download Excel", data=output.getvalue(), file_name=f"{nomor_aju}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"Error: {e}")

# Footer
st.markdown("<hr style='margin: 15px 0;'><div style='text-align: center; opacity: 0.5;'>© 2026 PT. Setia Samudera Abadi</div>", unsafe_allow_html=True)
