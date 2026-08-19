import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

# --- 1. PENGATURAN TAMPILAN UTAMA ---
st.set_page_config(
    page_title="KODEX - PT. Setia Samudera Abadi", page_icon="🚢", layout="wide"
)

# CSS Kustom untuk tampilan yang elegan, bersih, dan profesional
st.markdown(
    """
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
""",
    unsafe_allow_html=True,
)

# --- 2. HEADER APLIKASI & INPUT NOMOR AJU ---
st.title("KODEX 🚢")
st.markdown("### *Kompilator Dokumen Ekspor-Impor | PT. Setia Samudera Abadi*")

# Kotak input Nomor Aju untuk mengganti kolom NOMOR AJU di semua sheet Excel
nomor_aju = st.text_input(
    "🔢 Masukkan Nomor Aju", 
    placeholder="Contoh: 000020PRO32520260818000114"
)

st.write(
    "Unggah dokumen kepabeanan Anda. Sistem akan mengekstrak informasi dari"
    " Invoice, Packing List, HBL, MBL, dan Manifest BC 1.1 untuk disatukan"
    " menjadi satu file Excel template CEISA 4.0 dengan multi-sheet lengkap."
)
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
    st.info(
        "💡 Pastikan format tabel di dalam dokumen adalah versi final dan tidak"
        " dienkripsi."
    )

st.markdown("---")

# --- 4. TOMBOL PROSES & LOGIKA EKSTRAKSI ---
if st.button("🚀 Generate Excel CEISA 4.0", use_container_width=True):
    # Validasi apakah ada file yang diunggah dan nomor aju terisi
    if not (inv_file or pl_file or hbl_file or mbl_file or bc_file):
        st.warning("⚠️ Mohon unggah minimal satu dokumen untuk memulai proses.")
    elif not nomor_aju:
        st.warning("⚠️ Mohon isi Nomor Aju terlebih dahulu.")
    else:
        with st.spinner("Menganalisis dokumen dan menyusun matriks multi-sheet CEISA..."):
            try:
                items_data = []

                # [A] LOGIKA EKSTRAKSI INVOICE
                if inv_file:
                    with pdfplumber.open(inv_file) as pdf:
                        text = "\n".join(
                            [page.extract_text() for page in pdf.pages if page.extract_text()]
                        )

                    lines = text.split("\n")
                    current_item = {}
                    price_pattern = re.compile(r"^(\d+)\s+([\d\.]+)\s+([\d\.]+)$")

                    for i, line in enumerate(lines):
                        line = line.strip()
                        if "Prosind Code -" in line:
                            current_item["NOMOR AJU"] = nomor_aju
                            current_item["KODE BARANG"] = line.split("-")[-1].strip()
                            prev_line = lines[i - 1].strip()
                            match_desc_qty = re.match(r"^(.*?)\s+(\d+)$", prev_line)
                            if match_desc_qty:
                                current_item["URAIAN"] = match_desc_qty.group(1).strip().upper()
                                current_item["JUMLAH SATUAN"] = int(match_desc_qty.group(2))
                            else:
                                current_item["URAIAN"] = prev_line.upper()
                        elif "Net weight ea." in line:
                            weight_match = re.search(r"([\d\.]+)\s*kg", line)
                            if weight_match:
                                current_item["NETTO"] = float(weight_match.group(1))
                        elif "HS Code" in line:
                            hs_match = re.search(r"HS Code\s+([\d\.]+)", line)
                            if hs_match:
                                current_item["HS"] = hs_match.group(1).replace(".", "")
                        else:
                            price_match = price_pattern.match(line)
                            if price_match and "KODE BARANG" in current_item:
                                current_item["SERI BARANG"] = int(price_match.group(1))
                                current_item["HARGA SATUAN"] = float(price_match.group(2))
                                current_item["NILAI BARANG"] = current_item["HARGA SATUAN"]

                                # Data default template BARANG
                                current_item["MEREK"] = "PROSIND CONSULTING"
                                current_item["KODE KEMASAN"] = "BX"
                                current_item["METODE PENENTUAN NILAI"] = "Metode 1"

                                items_data.append(current_item)
                                current_item = {}

                # [B] PEMBUATAN STRUKTUR 21 MULTI-SHEET LENGKAP CEISA 4.0
                
                # 1. Sheet HEADER
                df_header = pd.DataFrame(columns=[
                    'NOMOR AJU', 'KODE DOKUMEN', 'KODE KANTOR', 'KODE KANTOR BONGKAR', 'KODE KANTOR PERIKSA', 
                    'KODE KANTOR TUJUAN', 'KODE KANTOR EKSPOR', 'KODE JENIS IMPOR', 'KODE JENIS EKSPOR', 'KODE JENIS TPB'
                ])
                if not df_header.empty or len(df_header.columns) > 0:
                    df_header.loc[0, 'NOMOR AJU'] = nomor_aju
                    df_header.loc[0, 'KODE DOKUMEN'] = 20
                    df_header.loc[0, 'KODE KANTOR'] = 50100

                # 2. Sheet ENTITAS
                df_entitas = pd.DataFrame(columns=[
                    'NOMOR AJU', 'SERI', 'KODE ENTITAS', 'KODE JENIS IDENTITAS', 'NOMOR IDENTITAS', 'NAMA ENTITAS'
                ])

                # 3. Sheet DOKUMEN
                df_dokumen = pd.DataFrame(columns=[
                    'NOMOR AJU', 'SERI', 'KODE DOKUMEN', 'NOMOR DOKUMEN', 'TANGGAL DOKUMEN', 'KODE FASILITAS', 'KODE IJIN'
                ])

                # 4. Sheet PENGANGKUT
                df_pengangkut = pd.DataFrame(columns=[
                    'NOMOR AJU', 'SERI', 'KODE CARA ANGKUT', 'NAMA PENGANGKUT', 'NOMOR PENGANGKUT', 'KODE BENDERA'
                ])

                # 5. Sheet KEMASAN
                df_kemasan = pd.DataFrame(columns=[
                    'NOMOR AJU', 'SERI', 'KODE KEMASAN', 'JUMLAH KEMASAN', 'MEREK', 'NOMOR SEGEL'
                ])

                # 6. Sheet KONTAINER
                df_kontainer = pd.DataFrame(columns=[
                    'NOMOR AJU', 'SERI', 'NOMOR KONTINER', 'KODE UKURAN KONTAINER', 'KODE JENIS KONTAINER', 'KODE TIPE KONTAINER', 'NOMOR SEGEL'
                ])

                # 7. Sheet KOMPONENBIAYA
                df_komponenbiaya = pd.DataFrame(columns=[
                    'NOMOR AJU', 'JENIS NILAI', 'HARGA INVOICE', 'PEMBAYARAN TIDAK LANGSUNG', 'DISKON'
                ])

                # 8. Sheet BARANG
                if items_data:
                    df_barang = pd.DataFrame(items_data)
                else:
                    df_barang = pd.DataFrame(columns=[
                        'NOMOR AJU', 'SERI BARANG', 'HS', 'KODE BARANG', 'URAIAN', 'MEREK', 'TIPE', 'UKURAN', 
                        'SPESIFIKASI LAIN', 'KODE SATUAN', 'JUMLAH SATUAN', 'KODE KEMASAN', 'JUMLAH KEMASAN', 
                        'NETTO', 'FOB', 'ASURANSI', 'FREIGHT', 'HARGA SATUAN', 'NILAI BARANG', 'METODE PENENTUAN NILAI'
                    ])

                # 9. Sheet BARANGTARIF
                df_barangtarif = pd.DataFrame(columns=[
                    'NOMOR AJU', 'SERI BARANG', 'KODE PUNGUTAN', 'KODE TARIF', 'TARIF', 'KODE FASILITAS'
                ])

                # 10. Sheet BARANGDOKUMEN
                df_barangdokumen = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI DOKUMEN', 'SERI IZIN'])

                # 11. Sheet BARANGENTITAS
                df_barangentitas = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI ENTITAS'])

                # 12. Sheet BARANGSPEKKHUSUS
                df_barangspekkhusus = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'KODE', 'URAIAN'])

                # 13. Sheet BARANGVD
                df_barangvd = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'KODE VD', 'NILAI BARANG'])

                # 14. Sheet BAHANBAKU
                df_bahanbaku = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI BAHAN BAKU', 'HS', 'KODE BARANG', 'URAIAN'])

                # 15. Sheet BAHANBAKUTARIF
                df_bahanbakutarif = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI BAHAN BAKU', 'KODE PUNGUTAN'])

                # 16. Sheet BAHANBAKUDOKUMEN
                df_bahanbakudokumen = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI BAHAN BAKU', 'SERI DOKUMEN'])

                # 17. Sheet PUNGUTAN
                df_pungutan = pd.DataFrame(columns=['NOMOR AJU', 'KODE FASILITAS TARIF', 'KODE JENIS PUNGUTAN', 'NILAI PUNGUTAN'])

                # 18. Sheet JAMINAN
                df_jaminan = pd.DataFrame(columns=['NOMOR AJU', 'KODE KANTOR', 'KODE JAMINAN', 'NOMOR JAMINAN'])

                # 19. Sheet BANKDEVISA
                df_bankdevisa = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'KODE', 'NAMA'])

                # 20. Sheet VERSI
                df_versi = pd.DataFrame({'VERSI': [1.3]})

                # 21. Sheet RESPON
                df_respon = pd.DataFrame(columns=['NOMOR AJU', 'KODE RESPON', 'NOMOR RESPON', 'TANGGAL RESPON'])

                # --- 5. MENULIS KE FILE EXCEL MULTI-SHEET ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_header.to_excel(writer, sheet_name="HEADER", index=False)
                    df_entitas.to_excel(writer, sheet_name="ENTITAS", index=False)
                    df_dokumen.to_excel(writer, sheet_name="DOKUMEN", index=False)
                    df_pengangkut.to_excel(writer, sheet_name="PENGANGKUT", index=False)
                    df_kemasan.to_excel(writer, sheet_name="KEMASAN", index=False)
                    df_kontainer.to_excel(writer, sheet_name="KONTAINER", index=False)
                    df_komponenbiaya.to_excel(writer, sheet_name="KOMPONENBIAYA", index=False)
                    df_barang.to_excel(writer, sheet_name="BARANG", index=False)
                    df_barangtarif.to_excel(writer, sheet_name="BARANGTARIF", index=False)
                    df_barangdokumen.to_excel(writer, sheet_name="BARANGDOKUMEN", index=False)
                    df_barangentitas.to_excel(writer, sheet_name="BARANGENTITAS", index=False)
                    df_barangspekkhusus.to_excel(writer, sheet_name="BARANGSPEKKHUSUS", index=False)
                    df_barangvd.to_excel(writer, sheet_name="BARANGVD", index=False)
                    df_bahanbaku.to_excel(writer, sheet_name="BAHANBAKU", index=False)
                    df_bahanbakutarif.to_excel(writer, sheet_name="BAHANBAKUTARIF", index=False)
                    df_bahanbakudokumen.to_excel(writer, sheet_name="BAHANBAKUDOKUMEN", index=False)
                    df_pungutan.to_excel(writer, sheet_name="PUNGUTAN", index=False)
                    df_jaminan.to_excel(writer, sheet_name="JAMINAN", index=False)
                    df_bankdevisa.to_excel(writer, sheet_name="BANKDEVISA", index=False)
                    df_versi.to_excel(writer, sheet_name="VERSI", index=False)
                    df_respon.to_excel(writer, sheet_name="RESPON", index=False)

                st.success(f"✅ File Excel multi-sheet CEISA 4.0 untuk Nomor Aju {nomor_aju} berhasil dibuat!")

                # Tombol download file Excel hasil ekstraksi
                st.download_button(
                    label="⬇️ Download Excel Format CEISA 4.0",
                    data=output.getvalue(),
                    file_name=f"{nomor_aju}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Terjadi kesalahan teknis saat membaca struktur dokumen: {e}")
