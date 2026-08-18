import streamlit as st
import pdfplumber
import pandas as pd
import io

# Membuat tata letak aplikasi lebih lebar
st.set_page_config(page_title="PDF ke Excel Converter", page_icon="📄", layout="wide")

st.title("Aplikasi Pengubah PDF ke Excel")
st.write("Unggah file PDF Anda. Jika tabel tidak terdeteksi, coba gunakan metode 'Tabel tanpa garis' di bawah.")

# Tambahkan pilihan untuk metode deteksi di bagian atas
st.subheader("Metode Deteksi Tabel")
detection_method = st.radio(
    "Pilih metode untuk file PDF Anda:",
    ("Tabel Bergaris Jelas (Lattice - Default)", "Tabel Tanpa Garis (Stream)"),
    help="Gunakan 'Lattice' untuk PDF dengan tabel kotak-kotak. Gunakan 'Stream' untuk PDF yang tabelnya dipisahkan spasi."
)

uploaded_file = st.file_uploader("Pilih file PDF", type="pdf")

if uploaded_file is not None:
    st.info("Membaca file PDF...")
    try:
        # Atur parameter pdfplumber berdasarkan pilihan
        pdfplumber_settings = {}
        if detection_method == "Tabel Tanpa Garis (Stream)":
            st.info("Mencoba deteksi tabel tanpa garis (Stream)...")
            # Atur tabel deteksi untuk pdfplumber 'stream'
            # Ini adalah tebakan terbaik untuk pengaturan stream umum
            pdfplumber_settings = {
                "vertical_strategy": "text", 
                "horizontal_strategy": "text", 
                "snap_tolerance": 3
            }
        else:
            st.info("Mencoba deteksi tabel bergaris default (Lattice)...")

        semua_tabel = []
        
        # Ekstrak data
        with pdfplumber.open(uploaded_file) as pdf:
            for i, page in enumerate(pdf.pages):
                # Gunakan pengaturan khusus jika dipilih
                tables = page.extract_tables(table_settings=pdfplumber_settings)
                for j, table in enumerate(tables):
                    if table:
                        semua_tabel.append((i+1, j+1, table))
        
        # Cek hasil dan proses Excel
        if len(semua_tabel) > 0:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for hal, tab_idx, table in semua_tabel:
                    # Buat DataFrame sederhana, biarkan user membersihkan di Excel
                    df = pd.DataFrame(table)
                    # Opsi: Jika Anda yakin baris pertama adalah header
                    # df = pd.DataFrame(table[1:], columns=table[0])
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
            if detection_method == "Tabel Tanpa Garis (Stream)":
                st.warning("Maaf, tabel masih tidak terdeteksi meskipun dengan metode 'Stream'. PDF ini mungkin merupakan hasil *scan* gambar, atau tata letaknya sangat rumit. Sistem ini tidak mendukung OCR (membaca teks dari gambar).")
            else:
                st.warning("Tabel tidak terdeteksi. Coba ubah metode deteksi menjadi 'Tabel Tanpa Garis (Stream)' di atas.")
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses: {e}")
