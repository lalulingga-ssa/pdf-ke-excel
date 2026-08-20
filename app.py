import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

# --- 1. PENGATURAN TAMPILAN UTAMA ---
st.set_page_config(
    page_title="KODEX - Kompilator Dokumen Ekspor-Impor", 
    page_icon="🚢", 
    layout="centered"
)

# CSS Kustom untuk tampilan yang bersih, rapi, dan pas satu layar (Compact Layout)
st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 900px;
    }
    h1 { font-size: 1.8rem !important; }
    h3 { font-size: 1.1rem !important; }
    
    /* Styling tombol utama agar compact dan menarik */
    .stButton>button {
        background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%);
        color: white; 
        border-radius: 6px;
        padding: 0.5rem 1rem; 
        border: none; 
        font-weight: bold;
        font-size: 15px;
        width: 100%;
        box-shadow: 0 2px 4px rgba(2, 132, 199, 0.2);
        transition: 0.3s;
    }
    .stButton>button:hover { 
        background: linear-gradient(135deg, #0369a1 100%, #075985 100%);
        color: white; 
    }
    </style>
""",
    unsafe_allow_html=True,
)

# --- 2. HEADER APLIKASI (Compact Header) ---
col_h1, col_h2 = st.columns([5, 1])
with col_h1:
    st.markdown("## **KODEX** 🚢 &nbsp; <span style='font-size: 15px; opacity: 0.7;'>Kompilator Dokumen Ekspor-Impor | PT. Setia Samudera Abadi</span>", unsafe_allow_html=True)
with col_h2:
    st.markdown("<div style='text-align: right; padding-top: 5px;'><a href='#' style='text-decoration: none; font-size: 13px;'>📖 Panduan</a></div>", unsafe_allow_html=True)

st.markdown("---")

# --- 3. PILIHAN JENIS DOKUMEN & INPUT PARAMETER ---
jenis_pemberitahuan = st.radio(
    "📌 Pilih Jenis Dokumen Kepabeanan:",
    ["PIB BC 2.0 (Impor)", "PEB BC 3.0 (Ekspor)"],
    horizontal=True
)

col_in1, col_in2 = st.columns(2)
with col_in1:
    nomor_aju = st.text_input(
        "🔢 Masukkan Nomor Aju", 
        placeholder="Contoh: 000020PRO32520260818000114"
    )
with col_in2:
    ndpbm_input = st.number_input(
        "💱 Masukkan Nilai NDPBM", 
        value=12644.5000, 
        format="%.4f"
    )

st.info("ℹ️ Unggah dokumen Invoice dan Packing List di bawah ini untuk langsung dikompilasi ke format CEISA 4.0.")

# --- 4. AREA UPLOAD DOKUMEN PENDUKUNG ---
st.subheader("📥 Unggah Dokumen Pendukung")

c1, c2, c3 = st.columns(3)
with c1:
    inv_file = st.file_uploader("1. Invoice (PDF)", type="pdf", key="inv")
with c2:
    pl_file = st.file_uploader("2. Packing List (PDF)", type="pdf", key="pl")
with c3:
    hbl_file = st.file_uploader("3. House B/L (PDF)", type="pdf", key="hbl")

c4, c5, c6 = st.columns(3)
with c4:
    mbl_file = st.file_uploader("4. Master B/L (PDF)", type="pdf", key="mbl")
with c5:
    bc_file = st.file_uploader("5. Manifest BC 1.1 (PDF)", type="pdf", key="bc")
with c6:
    st.markdown("<div style='padding-top: 28px; font-size: 12px; opacity: 0.8;'>💡 <b>Catatan:</b> Pastikan dokumen dalam format final dan tidak dienkripsi.</div>", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# --- 5. TOMBOL PROSES ---
generate_btn = st.button("🚀 Generate Excel CEISA 4.0", use_container_width=True)

# --- 6. LOGIKA EKSTRAKSI & GENERATE EXCEL ---
if generate_btn:
    if not inv_file:
        st.warning("⚠️ Mohon unggah dokumen Invoice terlebih dahulu untuk mengisi Sheet Barang.")
    elif not nomor_aju:
        st.warning("⚠️ Mohon isi Nomor Aju terlebih dahulu.")
    else:
        with st.spinner("Mengekstrak data PDF dan menyusun Sheet BARANG & HEADER..."):
            try:
                items_data = []
                merek_perusahaan = "PROSIND CONSULTING & ENGINEERING PTY.LTD."

                # [A] EKSTRAKSI TEKS PDF INVOICE
                inv_text = ""
                if inv_file:
                    with pdfplumber.open(inv_file) as pdf:
                        for page in pdf.pages:
                            p_text = page.extract_text()
                            if p_text:
                                inv_text += p_text + "\n"

                # [B] EKSTRAKSI TEKS PDF PACKING LIST (Untuk Netto)
                pl_text = ""
                netto_values = []
                if pl_file:
                    with pdfplumber.open(pl_file) as pdf:
                        for page in pdf.pages:
                            p_text = page.extract_text()
                            if p_text:
                                pl_text += p_text + "\n"
                    netto_matches = re.findall(r"Net weight ea\.?\s*([\d\.]+)\s*kg", pl_text, re.IGNORECASE)
                    if not netto_matches:
                        netto_matches = re.findall(r"Net weight\s*([\d\.]+)\s*kg", pl_text, re.IGNORECASE)
                    netto_values = [float(w) for w in netto_matches]

                # Ekstraksi baris dari Invoice berdasarkan penanda "Prosind Code"
                lines = inv_text.split("\n")
                current_item = {}
                seri_counter = 1

                for i, line in enumerate(lines):
                    line_clean = line.strip()
                    
                    if "Prosind Code" in line_clean:
                        if current_item:
                            items_data.append(current_item)
                            current_item = {}

                        current_item["NOMOR AJU"] = nomor_aju
                        current_item["SERI BARANG"] = seri_counter
                        seri_counter += 1
                        
                        # Ambil Kode Barang (6 digit angka)
                        code_match = re.search(r"(\d{6})", line_clean)
                        if code_match:
                            current_item["KODE BARANG"] = code_match.group(1)
                        else:
                            current_item["KODE BARANG"] = line_clean.split("-")[-1].strip()
                        
                        # Ambil Uraian & QTY secara akurat dari baris persis sebelumnya
                        if i > 0:
                            raw_prev = lines[i - 1].strip()
                            desc_match = re.match(r"^(\d+)(.+)$", raw_prev)
                            if desc_match:
                                current_item["JUMLAH SATUAN"] = int(desc_match.group(1))
                                current_item["URAIAN"] = desc_match.group(2).strip().upper()
                            else:
                                current_item["JUMLAH SATUAN"] = 1
                                current_item["URAIAN"] = raw_prev.upper()

                    elif "HS Code" in line_clean or "HS" in line_clean:
                        hs_match = re.search(r"(?:HS Code|HS)\s*[:\-]?\s*([\d\.]+)", line_clean, re.IGNORECASE)
                        if hs_match and current_item:
                            current_item["HS"] = hs_match.group(1).replace(".", "")
                        
                        price_hs_match = re.search(r"(\d+)\s+([\d\.]+)\s+([\d\.]+)$", line_clean)
                        if price_hs_match and current_item:
                            current_item["HARGA SATUAN"] = float(price_hs_match.group(2))
                            cif_val = float(price_hs_match.group(3))
                            current_item["CIF"] = cif_val
                            current_item["FOB"] = cif_val

                # Tangkap harga/CIF dari baris format reguler
                for i, line in enumerate(lines):
                    line_clean = line.strip()
                    price_pattern = re.compile(r"^(\d+)\s+([\d\.]+)\s+([\d\.]+)$")
                    price_match = price_pattern.match(line_clean)
                    if price_match and len(items_data) > 0:
                        seri_num = int(price_match.group(1))
                        for item in items_data:
                            if item.get("SERI BARANG") == seri_num:
                                item["HARGA SATUAN"] = float(price_match.group(2))
                                cif_val = float(price_match.group(3))
                                item["CIF"] = cif_val
                                item["FOB"] = cif_val

                if current_item:
                    items_data.append(current_item)

                # Mapping nilai default dan atribut tambahan ke setiap item barang
                for idx, item in enumerate(items_data):
                    if "HS" not in item:
                        item["HS"] = ""

                    if "URAIAN" not in item or not item["URAIAN"]:
                        item["URAIAN"] = f"BARANG IMPORT SERI {idx+1}"

                    if idx < len(netto_values):
                        item["NETTO"] = netto_values[idx]
                    else:
                        item["NETTO"] = 0.0

                    item["MEREK"] = merek_perusahaan
                    item["TIPE"] = "TANPA TIPE"
                    item["UKURAN"] = "-"
                    item["SPESIFIKASI LAIN"] = "-"
                    item["KODE SATUAN"] = "PCE"
                    item["KODE KEMASAN"] = "BX"
                    item["JUMLAH KEMASAN"] = 1
                    
                    if "CIF" not in item:
                        item["CIF"] = 0.0
                        item["FOB"] = 0.0

                    item["CIF RUPIAH"] = 89414570.19
                    item["NDPBM"] = ndpbm_input
                    item["ASURANSI"] = 0
                    item["FREIGHT"] = 247.42
                    item["NILAI TAMBAH"] = 0
                    item["DISKON"] = 0
                    item["HARGA PENYERAHAN"] = 0
                    item["HARGA PEROLEHAN"] = 0
                    if "HARGA SATUAN" not in item:
                        item["HARGA SATUAN"] = 0.0
                    item["HARGA EKSPOR"] = 0
                    item["NILAI BARANG"] = 0
                    item["NILAI JASA"] = 0
                    item["KODE JENIS NILAI"] = "LAI"
                    item["KODE KONDISI BARANG"] = 1
                    item["KODE NEGARA ASAL"] = "AU"
                    item["ISI PER KEMASAN"] = 0
                    item["METODE PENENTUAN NILAI"] = "Metode 1"
                    item["STATEMENT PERBEDAAN HARGA"] = "T"

                # [C] STRUKTUR MULTI-SHEET CEISA 4.0
                
                kode_dok_val = 20 if "PIB" in jenis_pemberitahuan else 30

                header_cols = [
                    'NOMOR AJU', 'KODE DOKUMEN', 'KODE KANTOR', 'KODE KANTOR BONGKAR', 'KODE KANTOR PERIKSA', 
                    'KODE KANTOR TUJUAN', 'KODE KANTOR EKSPOR', 'KODE JENIS IMPOR', 'KODE JENIS EKSPOR', 'KODE JENIS TPB', 
                    'KODE JENIS PLB', 'KODE JENIS PROSEDUR', 'KODE TUJUAN PEMASUKAN', 'KODE TUJUAN PENGIRIMAN', 'KODE TUJUAN TPB', 
                    'KODE CARA DAGANG', 'KODE CARA BAYAR', 'KODE CARA BAYAR LAINNYA', 'KODE GUDANG ASAL', 'KODE GUDANG TUJUAN', 
                    'KODE JENIS KIRIM', 'KODE JENIS PENGIRIMAN', 'KODE KATEGORI EKSPOR', 'KODE KATEGORI MASUK FTZ', 'KODE KATEGORI KELUAR FTZ', 
                    'KODE KATEGORI BARANG FTZ', 'KODE LOKASI', 'KODE LOKASI BAYAR', 'LOKASI ASAL', 'LOKASI TUJUAN', 
                    'KODE DAERAH ASAL', 'KODE GUDANG ASAL', 'KODE GUDANG TUJUAN', 'KODE NEGARA TUJUAN', 'KODE TUTUP PU', 
                    'NOMOR BC11', 'TANGGAL BC11', 'NOMOR POS', 'NOMOR SUB POS', 'KODE PELABUHAN BONGKAR', 
                    'KODE PELABUHAN MUAT', 'KODE PELABUHAN MUAT AKHIR', 'KODE PELABUHAN TRANSIT', 'KODE PELABUHAN TUJUAN', 'KODE PELABUHAN EKSPOR', 
                    'KODE TPS', 'TANGGAL BERANGKAT', 'TANGGAL EKSPOR', 'TANGGAL MASUK', 'TANGGAL MUAT', 
                    'TANGGAL TIBA', 'TANGGAL PERIKSA', 'TEMPAT STUFFING', 'TANGGAL STUFFING', 'KODE TANDA PENGAMAN', 
                    'JUMLAH TANDA PENGAMAN', 'FLAG CURAH', 'FLAG SDA', 'FLAG VD', 'FLAG AP BK', 
                    'FLAG MIGAS', 'KODE ASURANSI', 'ASURANSI', 'NILAI BARANG', 'NILAI INCOTERM', 
                    'NILAI MAKLON', 'ASURANSI', 'FREIGHT', 'FOB', 'BIAYA TAMBAHAN', 
                    'BIAYA PENGURANG', 'VD', 'CIF', 'HARGA_PENYERAHAN', 'NDPBM', 
                    'TOTAL DANA SAWIT', 'DASAR PENGENAAN PAJAK', 'NILAI JASA', 'UANG MUKA', 'BRUTO', 
                    'NETTO', 'VOLUME', 'KOTA PERNYATAAN', 'TANGGAL PERNYATAAN', 'NAMA PERNYATAAN', 
                    'JABATAN PERNYATAAN', 'KODE VALUTA', 'KODE INCOTERM', 'KODE JASA KENA PAJAK', 'NOMOR BUKTI BAYAR', 
                    'TANGGAL BUKTI BAYAR', 'KODE JENIS NILAI', 'KODE KANTOR MUAT', 'NOMOR DAFTAR', 'TANGGAL DAFTAR', 
                    'KODE ASAL BARANG FTZ', 'KODE TUJUAN PENGELUARAN', 'PPN PAJAK', 'PPNBM PAJAK', 'TARIF PPN PAJAK', 
                    'TARIF PPNBM PAJAK', 'BARANG TIDAK BERWUJUD', 'KODE JENIS PENGELUARAN', 'BARANG KIRIMAN', 'FLAG KONSOL', 
                    'KODE JENIS PENGANGKUTAN', 'FLAG PROPORSIONAL NETTO'
                ]
                
                df_header = pd.DataFrame(columns=header_cols)
                if not df_header.empty:
                    df_header.loc[0, 'NOMOR AJU'] = nomor_aju
                    df_header.loc[0, 'KODE DOKUMEN'] = kode_dok_val
                    df_header.loc[0, 'KODE KANTOR'] = 50100
                    df_header.loc[0, 'NDPBM'] = ndpbm_input

                # Sheet Entitas
                df_entitas = pd.DataFrame(columns=[
                    'NOMOR AJU', 'SERI', 'KODE ENTITAS', 'KODE JENIS IDENTITAS', 'NOMOR IDENTITAS', 
                    'NAMA ENTITAS', 'ALAMAT ENTITAS', 'NIB ENTITAS', 'KODE JENIS API', 'KODE STATUS', 
                    'NOMOR IJIN ENTITAS', 'TANGGAL IJIN ENTITAS', 'KODE NEGARA', 'NIPER ENTITAS', 
                    'KODE KATEGORI KONSOLIDATOR', 'KODE AFILIASI'
                ])
                
                df_dokumen = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'KODE DOKUMEN', 'NOMOR DOKUMEN', 'TANGGAL DOKUMEN', 'KODE FASILITAS', 'KODE IJIN'])
                df_pengangkut = pd.DataFrame(columns=[
                    'NOMOR AJU', 'SERI', 'KODE CARA ANGKUT', 'NAMA PENGANGKUT', 'NOMOR PENGANGKUT', 
                    'KODE BENDERA', 'CALL SIGN', 'FLAG ANGKUT PLB', 'CARA PENGANGKUTAN LAINNYA'
                ])
                df_kemasan = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'KODE KEMASAN', 'JUMLAH KEMASAN', 'MEREK', 'NOMOR SEGEL'])
                df_kontainer = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'NOMOR KONTINER', 'KODE UKURAN KONTAINER', 'KODE JENIS KONTAINER', 'KODE TIPE KONTAINER', 'NOMOR SEGEL'])
                df_komponenbiaya = pd.DataFrame(columns=[
                    'NOMOR AJU', 'JENIS NILAI', 'HARGA INVOICE', 'PEMBAYARAN TIDAK LANGSUNG', 'DISKON',
                    'KOMISI PENJUALAN', 'BIAYA PENGEMASAN', 'BIAYA PENGEPAKAN', 'ASSIST', 'ROYALTI',
                    'PROCEEDS', 'BIAYA TRANSPORTASI', 'BIAYA PEMUATAN', 'ASURANSI', 'GARANSI',
                    'BIAYA KEPENTINGAN SENDIRI', 'BIAYA PASCA IMPOR', 'BIAYA PAJAK INTERNAL', 'BUNGA', 'DEVIDEN'
                ])

                barang_columns = [
                    'NOMOR AJU', 'SERI BARANG', 'HS', 'KODE BARANG', 'URAIAN', 'MEREK', 'TIPE', 'UKURAN', 
                    'SPESIFIKASI LAIN', 'KODE SATUAN', 'JUMLAH SATUAN', 'KODE KEMASAN', 'JUMLAH KEMASAN', 
                    'KODE DOKUMEN ASAL', 'KODE KANTOR ASAL', 'NOMOR DAFTAR ASAL', 'TANGGAL DAFTAR ASAL', 
                    'NOMOR AJU ASAL', 'SERI BARANG ASAL', 'NETTO', 'BRUTO', 'VOLUME', 'SALDO AWAL', 
                    'SALDO AKHIR', 'JUMLAH REALISASI', 'CIF', 'CIF RUPIAH', 'NDPBM', 'FOB', 'ASURANSI', 
                    'FREIGHT', 'NILAI TAMBAH', 'DISKON', 'HARGA PENYERAHAN', 'HARGA PEROLEHAN', 
                    'HARGA SATUAN', 'HARGA EKSPOR', 'HARGA PATOKAN', 'NILAI BARANG', 'NILAI JASA', 
                    'NILAI DANA SAWIT', 'NILAI DEVISA', 'PERSENTASE IMPOR', 'KODE ASAL BARANG', 
                    'KODE DAERAH ASAL', 'KODE GUNA BARANG', 'KODE JENIS NILAI', 'JATUH TEMPO ROYALTI', 
                    'KODE KATEGORI BARANG', 'KODE KONDISI BARANG', 'KODE NEGARA ASAL', 'KODE PERHITUNGAN', 
                    'PERNYATAAN LARTAS', 'FLAG 4 TAHUN', 'SERI IZIN', 'TAHUN PEMBUATAN', 'KAPASITAS SILINDER', 
                    'KODE BKC', 'KODE KOMODITI BKC', 'KODE SUB KOMODITI BKC', 'FLAG TIS', 'ISI PER KEMASAN', 
                    'JUMLAH DILEKATKAN', 'JUMLAH PITA CUKAI', 'HJE CUKAI', 'TARIF CUKAI', 'KODE JENIS EKSPOR', 
                    'METODE PENENTUAN NILAI', 'ALASAN METODE PENENTUAN NILAI', 'STATEMENT PERBEDAAN HARGA'
                ]

                if items_data:
                    df_barang = pd.DataFrame(items_data)
                    for col in barang_columns:
                        if col not in df_barang.columns:
                            df_barang[col] = None
                    df_barang = df_barang[barang_columns]
                else:
                    df_barang = pd.DataFrame(columns=barang_columns)

                df_barangtarif = pd.DataFrame(columns=[
                    'NOMOR AJU', 'SERI BARANG', 'KODE PUNGUTAN', 'KODE TARIF', 'TARIF', 
                    'KODE FASILITAS', 'TARIF FASILITAS', 'NILAI BAYAR', 'NILAI FASILITAS', 
                    'NILAI SUDAH DILUNASI', 'KODE SATUAN', 'JUMLAH SATUAN', 'FLAG BMT SEMENTARA', 
                    'KODE KOMODITI CUKAI', 'KODE SUB KOMODITI CUKAI', 'FLAG TIS', 'FLAG PELEKATAN', 
                    'KODE KEMASAN', 'JUMLAH KEMASAN'
                ])
                
                df_barangdokumen = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI DOKUMEN', 'SERI IZIN'])
                df_barangentitas = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI ENTITAS'])
                df_barangspekkhusus = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'KODE', 'URAIAN'])
                df_barangvd = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'KODE VD', 'NILAI BARANG', 'BIAYA TAMBAHAN', 'BIAYA PENGURANG', 'JATUH TEMPO'])
                df_bahanbaku = pd.DataFrame(columns=[
                    'NOMOR AJU', 'SERI BARANG', 'SERI BAHAN BAKU', 'KODE ASAL BAHAN BAKU', 'HS', 'KODE BARANG', 
                    'URAIAN', 'MEREK', 'TIPE', 'UKURAN', 'SPESIFIKASI LAIN', 'KODE SATUAN', 'JUMLAH SATUAN', 
                    'KODE KEMASAN', 'JUMLAH KEMASAN', 'KODE DOKUMEN ASAL', 'KODE KANTOR ASAL', 'NOMOR DAFTAR ASAL', 
                    'TANGGAL DAFTAR ASAL', 'NOMOR AJU ASAL', 'SERI BARANG ASAL', 'NETTO', 'BRUTO', 'VOLUME', 
                    'CIF', 'CIF RUPIAH', 'NDPBM', 'HARGA PENYERAHAN', 'HARGA PEROLEHAN', 'NILAI JASA', 'SERI IZIN', 
                    'VALUTA', 'KODE BKC', 'KODE KOMODITI BKC', 'KODE SUB KOMODITI BKC', 'FLAG TIS', 'ISI PER KEMASAN', 
                    'JUMLAH DILEKATKAN', 'JUMLAH PITA CUKAI', 'HJE CUKAI', 'TARIF CUKAI'
                ])
                
                # --- UPDATE STRUKTUR SHEET BAHANBAKUTARIF: TERMASUK DUPLIKASI (KOLOM A S/D Y) ---
                df_bahanbakutarif = pd.DataFrame(columns=[
                    'NOMOR AJU', 'SERI BARANG', 'SERI BAHAN BAKU', 'KODE ASAL BAHAN BAKU', 
                    'KODE PUNGUTAN', 'KODE TARIF', 'TARIF', 'KODE FASILITAS', 'TARIF FASILITAS', 
                    'NILAI BAYAR', 'NILAI FASILITAS', 'NILAI SUDAH DILUNASI', 'KODE SATUAN', 
                    'JUMLAH SATUAN', 'FLAG BMT SEMENTARA', 'KODE KOMODITI CUKAI', 
                    'KODE SUB KOMODITI CUKAI', 'FLAG TIS', 'FLAG PELEKATAN', 'KODE KEMASAN', 
                    'KODE SUB KOMODITI CUKAI', 'FLAG TIS', 'FLAG PELEKATAN', 'KODE KEMASAN', 
                    'JUMLAH KEMASAN'
                ])
                
                df_bahanbakudokumen = pd.DataFrame(columns=[
                    'NOMOR AJU', 'SERI BARANG', 'SERI BAHAN BAKU', 'KODE_ASAL_BAHAN_BAKU', 
                    'SERI DOKUMEN', 'SERI IZIN'
                ])
                
                df_pungutan = pd.DataFrame(columns=[
                    'NOMOR AJU', 'KODE FASILITAS TARIF', 'KODE JENIS PUNGUTAN', 
                    'NILAI PUNGUTAN', 'NPWP BILLING'
                ])
                
                df_jaminan = pd.DataFrame(columns=[
                    'NOMOR AJU', 'KODE KANTOR', 'KODE JAMINAN', 'NOMOR JAMINAN', 'TANGGAL JAMINAN',
                    'NILAI JAMINAN', 'PENJAMIN', 'TANGGAL JATUH TEMPO', 'NOMOR BPJ', 'TANGGAL BPJ'
                ])
                
                df_bankdevisa = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'KODE', 'NAMA'])
                df_versi = pd.DataFrame({'VERSI': [1.3]})
                df_respon = pd.DataFrame(columns=['NOMOR AJU', 'KODE RESPON', 'NOMOR RESPON', 'TANGGAL RESPON'])

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

                st.success(f"✅ File Excel {jenis_pemberitahuan} untuk Nomor Aju {nomor_aju} berhasil di-generate!")

                st.download_button(
                    label=f"⬇️ Download Excel Format {jenis_pemberitahuan}",
                    data=output.getvalue(),
                    file_name=f"{nomor_aju}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Terjadi kesalahan teknis saat memproses dokumen: {e}")

# --- 7. FOOTER MINIMALIS ---
st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
st.markdown("<div style='text-align: center; font-size: 11px; opacity: 0.6;'>© 2026 PT. Setia Samudera Abadi &bull; v1.0.0</div>", unsafe_allow_html=True)
