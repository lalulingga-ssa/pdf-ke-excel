import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

# --- 1. PENGATURAN TAMPILAN UTAMA ---
st.set_page_config(
    page_title="KODEX - Kompilator Dokumen Ekspor-Impor", page_icon="🚢", layout="wide"
)

# CSS Kustom untuk menyamakan tampilan dashboard persis seperti desain referensi
st.markdown(
    """
    <style>
    .main { background-color: #f8fafc; }
    h1 { color: #0f172a; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-weight: 800; }
    
    /* Styling container kartu */
    .dashboard-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 24px;
        margin-bottom: 24px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    
    /* Tombol utama */
    .stButton>button {
        background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%);
        color: white; 
        border-radius: 8px;
        padding: 0.75rem 1rem; 
        border: none; 
        font-weight: bold;
        font-size: 16px;
        box-shadow: 0 4px 6px rgba(2, 132, 199, 0.2);
        transition: 0.3s;
    }
    .stButton>button:hover { 
        background: linear-gradient(135deg, #0369a1 0%, #075985 100%);
        color: white; 
    }
    
    /* Styling teks info */
    .footer-text {
        text-align: center;
        color: #64748b;
        font-size: 13px;
        margin-top: 40px;
        border-top: 1px solid #e2e8f0;
        padding-top: 20px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# --- 2. HEADER APLIKASI ---
col_head1, col_head2 = st.columns([6, 1])
with col_head1:
    st.markdown("## **KODEX** 🚢 &nbsp;&nbsp; <span style='font-size: 18px; color: #334155; font-weight: 600;'>Kompilator Dokumen Ekspor-Impor</span>", unsafe_allow_html=True)
    st.markdown("<span style='color: #64748b; font-size: 14px;'>PT. Setia Samudera Abadi</span>", unsafe_allow_html=True)
with col_head2:
    st.markdown("<div style='text-align: right; padding-top: 10px;'><a href='#' style='text-decoration: none; border: 1px solid #cbd5e1; padding: 6px 14px; border-radius: 8px; color: #334155; font-size: 13px;'>📖 Panduan</a></div>", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# --- 3. CONTAINER INPUT NOMOR AJU & NDPBM ---
with st.container():
    st.markdown("<div class='dashboard-card'>", unsafe_allow_html=True)
    
    col_in1, col_in2 = st.columns(2)
    with col_in1:
        st.markdown("🔢 **Masukkan Nomor Aju**")
        nomor_aju = st.text_input(
            "Nomor Aju",
            placeholder="Contoh: 000020PRO32520260818000114",
            label_visibility="collapsed"
        )
        st.markdown("<span style='color: #94a3b8; font-size: 12px;'>Contoh: 000020PRO32520260818000114</span>", unsafe_allow_html=True)
        
    with col_in2:
        st.markdown("💱 **Masukkan Nilai NDPBM**")
        ndpbm_input = st.number_input(
            "Nilai NDPBM",
            value=12644.5000,
            format="%.4f",
            label_visibility="collapsed"
        )
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.info("ℹ️ Unggah dokumen Invoice dan Packing List. Sistem akan mengekstrak informasi dan menyusunnya secara presisi ke dalam sheet BARANG dan multi-sheet CEISA 4.0 lainnya.")
    
    st.markdown("</div>", unsafe_allow_html=True)

# --- 4. AREA UPLOAD DOKUMEN PENDUKUNG ---
st.markdown("<div class='dashboard-card'>", unsafe_allow_html=True)
st.markdown("### 📥 Unggah Dokumen Pendukung")
st.markdown("<span style='color: #64748b; font-size: 14px;'>Format file yang didukung: PDF (Maksimal 200MB per file)</span>", unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

# Menggunakan 6 kolom untuk meniru tata letak kartu dokumen
d_col1, d_col2, d_col3, d_col4, d_col5, d_col6 = st.columns(6)

with d_col1:
    st.markdown("<div style='border: 1px dashed #cbd5e1; border-radius: 8px; padding: 12px; text-align: center; background: #fafafa;'>", unsafe_allow_html=True)
    st.markdown("<p style='font-size: 13px; font-weight: 600; color: #1e293b;'>1. Invoice (PDF)</p>", unsafe_allow_html=True)
    st.markdown("<div style='font-size: 36px;'>📄</div>", unsafe_allow_html=True)
    inv_file = st.file_uploader("Invoice", type="pdf", label_visibility="collapsed", key="inv")
    st.markdown("<span style='font-size: 11px; color: #94a3b8;'>200MB per file • PDF</span>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

with d_col2:
    st.markdown("<div style='border: 1px dashed #cbd5e1; border-radius: 8px; padding: 12px; text-align: center; background: #fafafa;'>", unsafe_allow_html=True)
    st.markdown("<p style='font-size: 13px; font-weight: 600; color: #1e293b;'>2. Packing List</p>", unsafe_allow_html=True)
    st.markdown("<div style='font-size: 36px;'>📄</div>", unsafe_allow_html=True)
    pl_file = st.file_uploader("Packing List", type="pdf", label_visibility="collapsed", key="pl")
    st.markdown("<span style='font-size: 11px; color: #94a3b8;'>200MB per file • PDF</span>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

with d_col3:
    st.markdown("<div style='border: 1px dashed #cbd5e1; border-radius: 8px; padding: 12px; text-align: center; background: #fafafa;'>", unsafe_allow_html=True)
    st.markdown("<p style='font-size: 13px; font-weight: 600; color: #1e293b;'>3. House B/L</p>", unsafe_allow_html=True)
    st.markdown("<div style='font-size: 36px;'>📄</div>", unsafe_allow_html=True)
    hbl_file = st.file_uploader("House B/L", type="pdf", label_visibility="collapsed", key="hbl")
    st.markdown("<span style='font-size: 11px; color: #94a3b8;'>200MB per file • PDF</span>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

with d_col4:
    st.markdown("<div style='border: 1px dashed #cbd5e1; border-radius: 8px; padding: 12px; text-align: center; background: #fafafa;'>", unsafe_allow_html=True)
    st.markdown("<p style='font-size: 13px; font-weight: 600; color: #1e293b;'>4. Master B/L</p>", unsafe_allow_html=True)
    st.markdown("<div style='font-size: 36px;'>📄</div>", unsafe_allow_html=True)
    mbl_file = st.file_uploader("Master B/L", type="pdf", label_visibility="collapsed", key="mbl")
    st.markdown("<span style='font-size: 11px; color: #94a3b8;'>200MB per file • PDF</span>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

with d_col5:
    st.markdown("<div style='border: 1px dashed #cbd5e1; border-radius: 8px; padding: 12px; text-align: center; background: #fafafa;'>", unsafe_allow_html=True)
    st.markdown("<p style='font-size: 13px; font-weight: 600; color: #1e293b;'>5. Manifest BC</p>", unsafe_allow_html=True)
    st.markdown("<div style='font-size: 36px;'>📄</div>", unsafe_allow_html=True)
    bc_file = st.file_uploader("Manifest BC", type="pdf", label_visibility="collapsed", key="bc")
    st.markdown("<span style='font-size: 11px; color: #94a3b8;'>200MB per file • PDF</span>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

with d_col6:
    st.markdown("<div style='border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px; text-align: center; background: #f8fafc; height: 100%; display: flex; flex-direction: column; justify-content: center;'>", unsafe_allow_html=True)
    st.markdown("<div style='font-size: 28px;'>💡</div>", unsafe_allow_html=True)
    st.markdown("<p style='font-size: 12px; color: #0284c7; font-weight: 500; margin-top: 8px;'>Pastikan dokumen dalam format final dan tidak dienkripsi.</p>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# Tombol Proses Full Width
generate_btn = st.button("🚀 Generate Excel CEISA 4.0", use_container_width=True)

st.markdown("</div>", unsafe_allow_html=True)

# --- 5. LOGIKA EKSTRAKSI & PROSES GENERATE EXCEL ---
if generate_btn:
    if not inv_file:
        st.warning("⚠️ Mohon unggah dokumen Invoice terlebih dahulu untuk mengisi Sheet Barang.")
    elif not nomor_aju:
        st.warning("⚠️ Mohon isi Nomor Aju terlebih dahulu.")
    else:
        with st.spinner("Mengekstrak data PDF (Invoice & Packing List) dan menyusun Sheet BARANG..."):
            try:
                items_data = []
                merek_perusahaan = "PROSIND CONSULTING & ENGINEERING PTY.LTD."

                # [A] EKSTRAKSI TEKS PDF INVOICE
                inv_text = ""
                if inv_file:
                    with pdfplumber.open(inv_file) as pdf:
                        inv_text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

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

                # Ekstraksi baris dari Invoice
                lines = inv_text.split("\n")
                current_item = {}
                seri_counter = 1

                for i, line in enumerate(lines):
                    line = line.strip()
                    if "Prosind Code -" in line:
                        if current_item:
                            items_data.append(current_item)
                            current_item = {}

                        current_item["NOMOR AJU"] = nomor_aju
                        current_item["SERI BARANG"] = seri_counter
                        seri_counter += 1
                        
                        current_item["KODE BARANG"] = line.split("-")[-1].strip()
                        
                        if i > 0:
                            prev_line = lines[i - 1].strip()
                            match_desc_qty = re.match(r"^(.*?)\s+(\d+)$", prev_line)
                            if match_desc_qty:
                                current_item["URAIAN"] = match_desc_qty.group(1).strip().upper()
                                current_item["JUMLAH SATUAN"] = int(match_desc_qty.group(2))
                            else:
                                current_item["URAIAN"] = prev_line.upper()
                                current_item["JUMLAH SATUAN"] = 1

                    elif "HS Code" in line:
                        hs_match = re.search(r"HS Code\s+([\d\.]+)", line)
                        if hs_match and current_item:
                            current_item["HS"] = hs_match.group(1).replace(".", "")
                        
                        price_hs_match = re.search(r"(\d+)\s+([\d\.]+)\s+([\d\.]+)$", line)
                        if price_hs_match and current_item:
                            current_item["HARGA SATUAN"] = float(price_hs_match.group(2))
                            cif_val = float(price_hs_match.group(3))
                            current_item["CIF"] = cif_val
                            current_item["FOB"] = cif_val

                for i, line in enumerate(lines):
                    line = line.strip()
                    price_pattern = re.compile(r"^(\d+)\s+([\d\.]+)\s+([\d\.]+)$")
                    price_match = price_pattern.match(line)
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

                for idx, item in enumerate(items_data):
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

                # [C] DEFINISI STRUKTUR 21 MULTI-SHEET CEISA 4.0
                df_header = pd.DataFrame(columns=[
                    'NOMOR AJU', 'KODE DOKUMEN', 'KODE KANTOR', 'KODE KANTOR BONGKAR', 'KODE KANTOR PERIKSA', 
                    'KODE KANTOR TUJUAN', 'KODE KANTOR EKSPOR', 'KODE JENIS IMPOR', 'KODE JENIS EKSPOR', 'KODE JENIS TPB'
                ])
                if not df_header.empty:
                    df_header.loc[0, 'NOMOR AJU'] = nomor_aju
                    df_header.loc[0, 'KODE DOKUMEN'] = 20
                    df_header.loc[0, 'KODE KANTOR'] = 50100

                df_entitas = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'KODE ENTITAS', 'KODE JENIS IDENTITAS', 'NOMOR IDENTITAS', 'NAMA ENTITAS'])
                df_dokumen = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'KODE DOKUMEN', 'NOMOR DOKUMEN', 'TANGGAL DOKUMEN', 'KODE FASILITAS', 'KODE IJIN'])
                df_pengangkut = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'KODE CARA ANGKUT', 'NAMA PENGANGKUT', 'NOMOR PENGANGKUT', 'KODE BENDERA'])
                df_kemasan = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'KODE KEMASAN', 'JUMLAH KEMASAN', 'MEREK', 'NOMOR SEGEL'])
                df_kontainer = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'NOMOR KONTINER', 'KODE UKURAN KONTAINER', 'KODE JENIS KONTAINER', 'KODE TIPE KONTAINER', 'NOMOR SEGEL'])
                df_komponenbiaya = pd.DataFrame(columns=['NOMOR AJU', 'JENIS NILAI', 'HARGA INVOICE', 'PEMBAYARAN TIDAK LANGSUNG', 'DISKON'])

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

                df_barangtarif = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'KODE PUNGUTAN', 'KODE TARIF', 'TARIF', 'KODE FASILITAS'])
                df_barangdokumen = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI DOKUMEN', 'SERI IZIN'])
                df_barangentitas = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI ENTITAS'])
                df_barangspekkhusus = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'KODE', 'URAIAN'])
                df_barangvd = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'KODE VD', 'NILAI BARANG'])
                df_bahanbaku = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI BAHAN BAKU', 'HS', 'KODE BARANG', 'URAIAN'])
                df_bahanbakutarif = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI BAHAN BAKU', 'KODE PUNGUTAN'])
                df_bahanbakudokumen = pd.DataFrame(columns=['NOMOR AJU', 'SERI BARANG', 'SERI BAHAN BAKU', 'SERI DOKUMEN'])
                df_pungutan = pd.DataFrame(columns=['NOMOR AJU', 'KODE FASILITAS TARIF', 'KODE JENIS PUNGUTAN', 'NILAI PUNGUTAN'])
                df_jaminan = pd.DataFrame(columns=['NOMOR AJU', 'KODE KANTOR', 'KODE JAMINAN', 'NOMOR JAMINAN'])
                df_bankdevisa = pd.DataFrame(columns=['NOMOR AJU', 'SERI', 'KODE', 'NAMA'])
                df_versi = pd.DataFrame({'VERSI': [1.3]})
                df_respon = pd.DataFrame(columns=['NOMOR AJU', 'KODE RESPON', 'NOMOR RESPON', 'TANGGAL RESPON'])

                # --- 6. MENULIS KE FILE EXCEL MULTI-SHEET ---
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

                st.success(f"✅ File Excel CEISA 4.0 untuk Nomor Aju {nomor_aju} berhasil di-generate!")

                st.download_button(
                    label="⬇️ Download Excel Format CEISA 4.0",
                    data=output.getvalue(),
                    file_name=f"{nomor_aju}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Terjadi kesalahan teknis saat memproses dokumen: {e}")

# --- 7. FOOTER ---
st.markdown("<div class='footer-text'>© 2026 PT. Setia Samudera Abadi. All rights reserved. &nbsp;&nbsp;|&nbsp;&nbsp; v1.0.0</div>", unsafe_allow_html=True)
