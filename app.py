import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(page_title="PDF ke Excel Converter", page_icon="📄")
st.title("Aplikasi Pengubah PDF ke Excel")
st.write("Unggah file PDF Anda, dan sistem akan mengekstrak tabelnya menjadi file Excel yang siap diunduh.")

uploaded_file = st.file_uploader("Pilih file PDF", type="pdf")

if uploaded_file is not None:
    st.info("Membaca file PDF...")
    try:
        semua_tabel = []
        
        # 1. Ekstrak data dulu tanpa membuat Excel
        with pdfplumber.open(uploaded_file) as pdf:
            for i, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                for j, table in enumerate(tables):
                    if table:
                        semua_tabel.append((i+1, j+1, table))
        
        # 2. Cek apakah ada tabel yang berhasil diekstrak
        if len(semua_tabel) > 0:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for hal, tab_idx, table in semua_tabel:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    sheet_name = f"Hal_{hal}_Tab_{tab_idx}"
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            st.success(f"Berhasil! Ditemukan {len(semua_tabel)} tabel.")
            st.download_button(
                label="⬇️ Download File Excel",
                data=output.getvalue(),
                file_name="Hasil_Konversi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Tabel tidak terdeteksi. Sistem ini paling optimal untuk tabel PDF yang memiliki garis kotak (border) yang jelas.")
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses: {e}")
