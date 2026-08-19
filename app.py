import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# --- 1. PENGATURAN TAMPILAN UTAMA ---
st.set_page_config(page_title="OptiCEISA DataForge", page_icon="🚢", layout="wide")

# CSS Kustom untuk tampilan yang elegan, bersih, dan profesional
st.markdown("""
    <style>
    .main { background-color: #f4f6f9; }
    h1 { color: #0a3d62; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-weight: 800; }
    h3 { color: #3c6382; }
    .stButton>button {
        background-color: #0a3d62; color: white; border-radius: 6px;
        padding: 0.6rem 1rem; border: none; transition: 0.3s; font-weight: bold;
    }
    .stButton>button:hover { background-color: #38ada9; color: white; border: none; }
    </style>
""", unsafe_allow_html=True)

# --- 2. HEADER APLIKASI ---
st.title("OptiCEISA DataForge ⚡")
st.markdown("### *Intelligent CEISA 4.0 Document Extractor & Integrator*")
st.write("Unggah dokumen kepabeanan Anda. Sistem akan mengekstrak informasi dari Invoice, Packing List, HBL, MBL, dan Manifest BC 1.1 untuk disatukan menjadi satu file Excel template CEISA 4.0.")
st.markdown("---")

# --- 3. AREA UPLOAD (Dibuat grid/kolom agar rapi) ---
st.subheader("📥 Unggah Dokumen Pendukung")

# Baris pertama (3 Kolom)
col1, col2, col3 = st.columns(3)
with col1:
    inv_file = st.file_uploader("1. Invoice", type="pdf")
with col2:
    pl_file = st.file_uploader("2. Packing List", type="pdf")
with col3:
    hbl_file = st.file_uploader("3. House B/L", type="pdf")

# Baris kedua (2 Kolom + 1 Info)
col4, col5, col6 = st.columns(3)
with col4:
    mbl_file = st.file_uploader("4. Master B/L", type="pdf")
with col5:
    bc_file = st.file_uploader("5. Manifest BC 1.1", type="pdf")
with col6:
    st.info("💡 Pastikan format tabel di dalam dokumen adalah versi final dan tidak dienkripsi.")

st.markdown("---")

# --- 4. TOMBOL PROSES & LOGIKA EKSTRAKSI ---
if st.button("🚀 Generate Excel CEISA 4.0", use_container_width=True):
    # Validasi apakah ada file yang diunggah
    if not (inv_file or pl_file or hbl_file or mbl_file or bc_file):
        st.warning("⚠️ Mohon unggah minimal satu dokumen untuk memulai proses.")
    else:
        with st.spinner("Menganalisis dokumen dan menyusun matriks data CEISA..."):
            try:
                items_data = []
                
                # [A] LOGIKA EKSTRAKSI INVOICE (Mempertahankan logika Proshind sebelumnya)
                if inv_file:
                    with pdfplumber.open(inv_file) as pdf:
                        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
                    
                    lines = text.split('\n')
                    current_item = {}
                    price_pattern = re.compile(r'^(\d+)\s+([\d\.]+)\s+([\d\.]+)$')
                    
                    for i, line in enumerate(lines):
                        line = line.strip()
                        if "Prosind Code -" in line:
                            current_item['KODE BARANG'] = line.split("-")[-1].strip()
                            prev_line = lines[i-1].strip()
                            match_desc_qty = re.match(r'^(.*?)\s+(\d+)$', prev_line)
                            if match_desc_qty:
                                current_item['URAIAN'] = match_desc_qty.group(1).strip().upper()
                                current_item['JUMLAH SATUAN'] = int(match_desc_qty.group(2))
                            else:
                                current_item['URAIAN'] = prev_line.upper()
                        elif "Net weight ea." in line:
                            weight_match = re.search(r'([\d\.]+)\s*kg', line)
                            if weight_match:
                                current_item['NETTO'] = float(weight_match.group(1))
                        elif "HS Code" in line:
                            hs_match = re.search(r'HS Code\s+([\d\.]+)', line)
                            if hs_match:
                                current_item['HS'] = hs_match.group(1).replace(".", "")
                        else:
                            price_match = price_pattern.match(line)
                            if price_match and 'KODE BARANG' in current_item:
                                current_item['SERI BARANG'] = int(price_match.group(1))
                                current_item['HARGA SATUAN'] = float(price_match.group(2))
                                current_item['NILAI BARANG'] = current_item['HARGA SATUAN']
                                
                                # Data default template BARANG
                                current_item['MEREK'] = "PROSIND CONSULTING"
                                current_item['KODE KEMASAN'] = "BX"
                                current_item['METODE PENENTUAN NILAI'] = "Metode 1"
                                
                                items_data.append(current_item)
                                current_item = {}

                # [B] TEMPAT UNTUK LOGIKA PL, HBL, MBL, BC 1.1
                # Catatan: Logika regex untuk file selain Invoice akan ditambahkan
                # setelah format/struktur teks masing-masing PDF diketahui.
                
                # --- 5. PEMBUATAN FILE EXCEL MULTI-SHEET ---
                # Membuat DataFrame untuk Sheet BARANG (Dari Invoice & PL)
                if items_data:
                    df_barang = pd.DataFrame(items_data)
                else:
                    df_barang = pd.DataFrame(columns=['SERI BARANG', 'KODE BARANG', 'URAIAN', 'NETTO', 'JUMLAH SATUAN', 'HARGA SATUAN'])

                # Membuat DataFrame untuk Sheet DOKUMEN (Dari HBL, MBL, BC 1.1, Invoice, PL)
                df_dokumen = pd.DataFrame(columns=[
                    'SERI DOKUMEN', 'KODE DOKUMEN', 'NOMOR DOKUMEN', 'TANGGAL DOKUMEN'
                ])
                
                # Membuat DataFrame untuk Sheet HEADER
                df_header = pd.DataFrame(columns=[
                    'NOMOR PENGAJUAN', 'JALUR', 'KODE ASAL BARANG', 'NILAI INVOICE', 'ASURANSI', 'FREIGHT'
                ])

                # Proses menyatukan Sheet ke dalam satu file Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_header.to_excel(writer, sheet_name='HEADER', index=False)
                    df_dokumen.to_excel(writer, sheet_name='DOKUMEN', index=False)
                    df_barang.to_excel(writer, sheet_name='BARANG', index=False)
                
                st.success("✅ Seluruh data berhasil diproses dan dikompilasi!")
                
                st.download_button(
                    label="⬇️ Download Excel Hasil Ekstraksi (Format CEISA)",
                    data=output.getvalue(),
                    file_name="OptiCEISA_Data_Kompilasi.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Terjadi kesalahan teknis saat membaca struktur dokumen: {e}")
