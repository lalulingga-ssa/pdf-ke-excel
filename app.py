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
        # Siapkan tempat untuk menyimpan Excel di memori
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            with pdfplumber.open(uploaded_file) as pdf:
                tabel_ditemukan = 0
                for i, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    for j, table in enumerate(tables):
                        if table:
                            tabel_ditemukan += 1
                            # Baris pertama PDF jadi header Excel
                            df = pd.DataFrame(table[1:], columns=table[0])
                            sheet_name = f"Hal_{i+1}_Tab_{j+1}"
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        if tabel_ditemukan > 0:
            st.success(f"Berhasil! Ditemukan {tabel_ditemukan} tabel.")
            
            # Tombol Download
            st.download_button(
                label="⬇️ Download File Excel",
                data=output.getvalue(),
                file_name="Hasil_Konversi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Tidak ada tabel bergaris yang terdeteksi di dalam file PDF ini.")
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses: {e}")
