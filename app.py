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
    # Default kolom
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
    st.subheader("Tambah Data Customer")
    tipe_cust = st.radio("Pilih Tipe:", ["Importir", "Eksportir"], horizontal=True)
    
    with st.form("customer_form"):
        if tipe_cust == "Importir":
            c1, c2 = st.columns(2)
            with c1:
                st.write("**1. Pengirim**")
                p_nama = st.text_input("Nama Pengirim")
                p_alamat = st.text_input("Alamat Pengirim")
                p_bendera = st.text_input("Bendera")
                st.write("**2. Penjual**")
                j_nama = st.text_input("Nama Penjual")
                j_alamat = st.text_input("Alamat Penjual")
                j_bendera = st.text_input("Bendera Penjual")
            with c2:
                st.write("**3. Importir**")
                i_nama = st.text_input("Nama Importir")
                i_alamat = st.text_input("Alamat Importir")
                i_id = st.text_input("No Identitas")
                i_nib = st.text_input("No NIB")
                i_nitku = st.text_input("No NITKU")
                st.write("**4. Pemilik Barang**")
                pb_nama = st.text_input("Nama Pemilik")
                pb_alamat = st.text_input("Alamat Pemilik")
                pb_nitku = st.text_input("No NITKU Pemilik")
                pb_relasi = st.text_input("Hubungan dengan Penjual")
        
        if st.form_submit_button("💾 Simpan Data Baru"):
            new_row = pd.DataFrame([[tipe_cust, p_nama, j_nama, i_nama, pb_nama]], columns=["Tipe", "Pengirim", "Penjual", "Importir", "Pemilik_Barang"])
            save_db(pd.concat([load_db(), new_row], ignore_index=True))
            st.success("Data berhasil disimpan!")

    st.divider()
    st.subheader("Edit/Hapus Database")
    
    # Fitur Edit dan Hapus menggunakan Data Editor
    df_db = load_db()
    edited_df = st.data_editor(df_db, num_rows="dynamic", use_container_width=True)
    
    if st.button("🔄 Simpan Perubahan Tabel (Hapus/Edit)"):
        save_db(edited_df)
        st.success("Perubahan pada database telah disimpan!")

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
                df_entitas = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'KODE ENTITAS', 'KODE JENIS IDENTITAS'])
                # Data Entitas baris 4-7
                data_entitas = {'NOMOR AJU': [nomor_aju]*7, 'SERI': [10, 9, 11, 1, 7, 4, 10], 'KODE ENTITAS': [10, 9, 11, 1, 7, 4, 10], 'KODE JENIS IDENTITAS': [None, None, None, 6, 6, 6, 6]}
                df_entitas = pd.concat([df_entitas, pd.DataFrame(data_entitas)], ignore_index=True)

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
