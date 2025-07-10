# STREAMLIT APLIKASI PENILAIAN PENGURUS BUMDes DENGAN REKAP & WORD FINAL

import streamlit as st
import pandas as pd
import os
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import qrcode
from io import BytesIO

# --- Path data
DATA_FOLDER = "data"
HASIL_FILE = f"{DATA_FOLDER}/hasil_penilaian.csv"
KANDIDAT_FILE = f"{DATA_FOLDER}/kandidat.csv"
os.makedirs(DATA_FOLDER, exist_ok=True)

# --- Load kandidat
if os.path.exists(KANDIDAT_FILE):
    kandidat_df = pd.read_csv(KANDIDAT_FILE)
else:
    kandidat_df = pd.DataFrame({"Nama": [], "Posisi": []})

# --- Form Identitas Penilai
st.subheader("🧑‍⚖️ Identitas Penilai")
if "penilai_info" not in st.session_state:
    with st.form("form_penilai"):
        nama_penilai = st.text_input("Nama Lengkap")
        jabatan_penilai = st.text_input("Jabatan")
        lembaga_penilai = st.text_input("Asal Lembaga (Pemdes/BPD/Kecamatan/DPMPD)")
        submit_id = st.form_submit_button("✅ Simpan Identitas")
    if submit_id and nama_penilai and jabatan_penilai and lembaga_penilai:
        st.session_state.penilai_info = {
            "nama": nama_penilai.strip(),
            "jabatan": jabatan_penilai.strip(),
            "lembaga": lembaga_penilai.strip()
        }
        st.success("Identitas penilai disimpan.")
    st.stop()
else:
    penilai = st.session_state.penilai_info
    st.success(f"Penilai: {penilai['nama']} ({penilai['jabatan']} - {penilai['lembaga']})")

# --- Filter kandidat yg belum dinilai oleh penilai ini
penilai_nama = penilai["nama"]
kandidat_tersedia = kandidat_df.copy()
if os.path.exists(HASIL_FILE):
    hasil_df = pd.read_csv(HASIL_FILE)
    sudah_dinilai = hasil_df[hasil_df["Nama Penilai"] == penilai_nama][["Nama", "Posisi"]]
    kandidat_tersedia = pd.merge(kandidat_df, sudah_dinilai, on=["Nama", "Posisi"], how="left", indicator=True)
    kandidat_tersedia = kandidat_tersedia[kandidat_tersedia["_merge"] == "left_only"].drop(columns=["_merge"])

if kandidat_tersedia.empty:
    st.info("✅ Anda telah menilai semua kandidat. Terima kasih.")
    st.stop()

# --- Pilih posisi dan kandidat
st.subheader("📝 Form Penilaian")
posisi_pilih = st.selectbox("Pilih Posisi yang Dinilai:", kandidat_tersedia["Posisi"].unique())
kandidat_pilih = st.selectbox("Pilih Kandidat:", kandidat_tersedia[kandidat_tersedia["Posisi"] == posisi_pilih]["Nama"].unique())

# --- Form penilaian skor
with st.form("form_penilaian"):
    psikologi = st.number_input("Tes Psikologi (0-100)", 0, 100, value=0)
    office = st.number_input("Tes Microsoft Office (0-100)", 0, 100, value=0)
    presentasi = st.number_input("Presentasi Gagasan (0-100)", 0, 100, value=0)
    esai = st.number_input("Esai & Refleksi Diri (0-100)", 0, 100, value=0)
    wawancara = st.number_input("Wawancara Panel (0-100)", 0, 100, value=0)
    simpan_nilai = st.form_submit_button("💾 Simpan Penilaian")

if simpan_nilai:
    nilai_baru = pd.DataFrame([{
        "Nama": kandidat_pilih,
        "Posisi": posisi_pilih,
        "Nama Penilai": penilai["nama"],
        "Jabatan": penilai["jabatan"],
        "Lembaga": penilai["lembaga"],
        "Tes Psikologi": psikologi,
        "Tes MS Office": office,
        "Presentasi Gagasan": presentasi,
        "Esai Refleksi Diri": esai,
        "Wawancara Panel": wawancara
    }])
    if os.path.exists(HASIL_FILE):
        hasil_lama = pd.read_csv(HASIL_FILE)
        hasil_df = pd.concat([hasil_lama, nilai_baru], ignore_index=True)
    else:
        hasil_df = nilai_baru
    hasil_df.to_csv(HASIL_FILE, index=False)
    st.success("Penilaian berhasil disimpan. Silakan lanjut menilai posisi/kandidat lainnya.")
    st.experimental_rerun()

# --- Export Word Laporan
st.subheader("📄 Export Rekap Penilaian (Word)")
if st.button("📥 Generate Word Rekap"):
    if not os.path.exists(HASIL_FILE):
        st.warning("Belum ada data penilaian.")
    else:
        df = pd.read_csv(HASIL_FILE)
        df["Total"] = df[["Tes Psikologi", "Tes MS Office", "Presentasi Gagasan", "Esai Refleksi Diri", "Wawancara Panel"]].apply(
            lambda r: r["Tes Psikologi"]*0.15 + r["Tes MS Office"]*0.15 + r["Presentasi Gagasan"]*0.3 + r["Esai Refleksi Diri"]*0.2 + r["Wawancara Panel"]*0.2,
            axis=1
        )
        ranking = df.groupby(["Nama", "Posisi"]).agg({"Total":"mean"}).reset_index()
        ranking = ranking.sort_values(["Posisi", "Total"], ascending=[True, False])

        doc = Document()
        title = doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title.add_run("LAPORAN HASIL PENILAIAN\nPENGURUS BUMDes Buwana Raharja Desa Keling")
        title_run.bold = True
        title_run.font.size = Pt(14)

        for posisi in ranking["Posisi"].unique():
            doc.add_paragraph("\n")
            doc.add_paragraph(f"🏆 {posisi}").runs[0].bold = True
            posisi_df = ranking[ranking["Posisi"] == posisi].reset_index(drop=True)
            table = doc.add_table(rows=1, cols=4)
            hdr = table.rows[0].cells
            hdr[0].text = "No"
            hdr[1].text = "Nama"
            hdr[2].text = "Total Skor"
            hdr[3].text = "Penghargaan"
            for i, row in posisi_df.iterrows():
                cells = table.add_row().cells
                cells[0].text = str(i+1)
                cells[1].text = row["Nama"]
                cells[2].text = f"{row['Total']:.2f}"
                cells[3].text = ["🥇 Juara 1 (Emas)", "🥈 Juara 2 (Perak)", "🥉 Juara 3 (Perunggu)"] [i] if i < 3 else "-"
            doc.add_paragraph(f"🎉 Selamat kepada {posisi_df.iloc[0]['Nama']} atas pencapaian nilai tertinggi dan ditetapkan sebagai calon terbaik {posisi}.")

        doc.add_paragraph("\n\nLembar Pengesahan Penilai:").runs[0].bold = True
        table = doc.add_table(rows=1, cols=3)
        table.rows[0].cells[0].text = "Nama Penilai"
        table.rows[0].cells[1].text = "Jabatan"
        table.rows[0].cells[2].text = "Tanda Tangan"
        penilai_unik = df[["Nama Penilai", "Jabatan"]].drop_duplicates()
        for _, row in penilai_unik.iterrows():
            row_ = table.add_row().cells
            row_[0].text = row["Nama Penilai"]
            row_[1].text = row["Jabatan"]
            row_[2].text = ".............................."

        qr = qrcode.make("Dokumen sah - Panitia Pemilihan BUMDes Desa Keling")
        buf = BytesIO()
        qr.save(buf)
        buf.seek(0)
        doc.add_picture(buf, width=Inches(1.5))
        doc.add_paragraph("Barcode ini menunjukkan dokumen resmi yang diterbitkan oleh Panitia Pemilihan Pengurus BUMDes Buwana Raharja Desa Keling.")

        path = f"{DATA_FOLDER}/Rekap_Final_Penilaian_BUMDes.docx"
        doc.save(path)
        with open(path, "rb") as f:
            st.download_button("📄 Download Word", f, file_name="Rekap_Penilaian_BUMDes.docx")

# Footer
st.markdown("<div style='text-align:center'>Developed by CV Mitra Utama Consultindo</div>", unsafe_allow_html=True)
