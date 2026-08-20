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
    return pd.DataFrame(columns=["Tipe", "Nama", "Alamat", "No_Identitas", "NIB", "NITKU", "Info_Tambahan"])

def save_db(df):
    df.to_csv(DB_FILE, index=False)

# --- 4. STATE MANAGEMENT ---
if 'edit_index' not in st.session_state: st.session_state.edit_index = None

# --- 5. TABS ---
tab1, tab2, tab3, tab4 = st.tabs(["📌 Data Utama", "📥 Dokumen", "🗄️ Database Customer", "🚀 Proses"])

with tab1:
    jenis_pemberitahuan = st.radio("Pilih Jenis Dokumen:", ["PIB BC 2.0 (Impor)", "PEB BC 3.0 (Ekspor)"], horizontal=True)
    col1, col2 = st.columns(2)
    nomor_aju = col1.text_input("🔢 Nomor Aju")
    ndpbm_input = col2.number_input("💱 Nilai NDPBM", value=12644.5000, format="%.4f")

with tab2:
    inv_file = st.file_uploader("1. Invoice", type="pdf")
    pl_file = st.file_uploader("2. Packing List", type="pdf")

with tab3:
    st.subheader("Manajemen Customer")
    
    # Form Tambah/Edit
    with st.form("customer_form"):
        tipe_cust = st.radio("Pilih Tipe:", ["Importir", "Eksportir"], horizontal=True)
        c1, c2 = st.columns(2)
        nama = c1.text_input("Nama (Pengirim/Importir/etc)")
        alamat = c2.text_input("Alamat")
        c3, c4, c5 = st.columns(3)
        no_id = c3.text_input("No Identitas")
        nib = c4.text_input("NIB")
        nitku = c5.text_input("NITKU")
        
        submitted = st.form_submit_button("💾 Simpan Data")
        if submitted:
            new_data = pd.DataFrame([[tipe_cust, nama, alamat, no_id, nib, nitku, ""]], columns=["Tipe", "Nama", "Alamat", "No_Identitas", "NIB", "NITKU", "Info_Tambahan"])
            save_db(pd.concat([load_db(), new_data], ignore_index=True))
            st.success("Data disimpan!")

    st.divider()
    
    # Menampilkan Tabel dengan Tombol Edit & Hapus
    df = load_db()
    for index, row in df.iterrows():
        cols = st.columns([4, 1, 1]) # Layout kolom data | Edit | Hapus
        cols[0].write(f"**{row['Tipe']}** - {row['Nama']} ({row['Alamat']})")
        
        if cols[1].button("✏️", key=f"edit_{index}"):
            st.session_state.edit_index = index
            st.rerun()
        if cols[2].button("🗑️", key=f"del_{index}"):
            save_db(df.drop(index))
            st.rerun()

with tab4:
    if st.button("🚀 Generate Excel", use_container_width=True):
        st.write("Proses ekstraksi dijalankan...")

# Footer
st.markdown("<hr style='margin: 15px 0;'><div style='text-align: center; opacity: 0.5;'>© 2026 PT. Setia Samudera Abadi</div>", unsafe_allow_html=True)
