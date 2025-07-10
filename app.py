# STREAMLIT APLIKASI PENILAIAN PENGURUS BUMDes DENGAN REKAP & WORD

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
PENILAI_FILE = f"{DATA_FOLDER}/penilai.csv"
KANDIDAT_FILE = f"{DATA_FOLDER}/kandidat.csv"
os.makedirs(DATA_FOLDER, exist_ok=True)

# --- Load kandidat dan penilai
if os.path.exists(KANDIDAT_FILE):
    kandidat_df = pd.read_csv(KANDIDAT_FILE)
else:
    kandidat_df = pd.DataFrame({"Nama": [], "Posisi": []})

if os.path.exists(PENILAI_FILE):
    penilai_df = pd.read_csv(PENILAI_FILE)
else:
    penilai_df = pd.DataFrame({"Nama Penilai": []})

# --- Form identitas penilai
st.subheader("🧑‍⚖️ Identitas Penilai")
nama_penilai = st.selectbox("Pilih Identitas Penilai:", penilai_df["Nama Penilai"].unique())

# --- Form penilaian
st.subheader("📝 Form Penilaian")
posisi_pilih = st.selectbox("Pilih Posisi yang Dinilai:", kandidat_df["Posisi"].unique())
kandidat_pilih = st.selectbox("Pilih Kandidat:", kandidat_df[kandidat_df["Posisi"] == posisi_pilih]["Nama"].unique())

with st.form("form_penilaian"):
    psikologi = st.number_input("Tes Psikologi (0-100)", 0, 100)
    office = st.number_input("Tes Microsoft Office (0-100)", 0, 100)
    presentasi = st.number_input("Presentasi Gagasan (0-100)", 0, 100)
    esai = st.number_input("Esai & Refleksi Diri (0-100)", 0, 100)
    wawancara = st.number_input("Wawancara Panel (0-100)", 0, 100)
    submit = st.form_submit_button("💾 Simpan Penilaian")

if submit:
    nilai_baru = pd.DataFrame.from_dict([{
        "Nama": kandidat_pilih,
        "Posisi": posisi_pilih,
        "Nama Penilai": nama_penilai,
        "Tes Psikologi": psikologi,
        "Tes MS Office": office,
        "Presentasi Gagasan": presentasi,
        "Esai Refleksi Diri": esai,
        "Wawancara Panel": wawancara
    }])

    if os.path.exists(HASIL_FILE):
        hasil_lama = pd.read_csv(HASIL_FILE)
        # Cegah duplikasi penilai terhadap kandidat yang sama
        hasil_lama = hasil_lama[~((hasil_lama["Nama"] == kandidat_pilih) & 
                                  (hasil_lama["Nama Penilai"] == nama_penilai))]
        hasil_df = pd.concat([hasil_lama, nilai_baru], ignore_index=True)
    else:
        hasil_df = nilai_baru

    hasil_df.to_csv(HASIL_FILE, index=False)
    st.success("✅ Penilaian berhasil disimpan.")

# --- Export Word dengan piala, ucapan, QR dan pengesahan
st.subheader("📄 Export Word Rekap dengan Piala dan Ucapan")
if st.button("📥 Generate Word Laporan"):
    hasil_df = pd.read_csv(HASIL_FILE)
    hasil_df["Total"] = hasil_df[[
        "Tes Psikologi", "Tes MS Office", "Presentasi Gagasan", "Esai Refleksi Diri", "Wawancara Panel"
    ]].apply(lambda r: r["Tes Psikologi"] * 0.15 + r["Tes MS Office"] * 0.15 + r["Presentasi Gagasan"] * 0.3 + r["Esai Refleksi Diri"] * 0.2 + r["Wawancara Panel"] * 0.2, axis=1)

    ranking_df = hasil_df.groupby(["Nama", "Posisi"]).agg({"Total": "mean"}).reset_index()
    ranking_df = ranking_df.sort_values(["Posisi", "Total"], ascending=[True, False])

    doc = Document()

    # Judul
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("LAPORAN HASIL PENILAIAN\nPENGURUS BUMDes Buwana Raharja Desa Keling")
    run.bold = True
    run.font.size = Pt(14)

    posisi_list = ranking_df["Posisi"].unique()
    for posisi in posisi_list:
        doc.add_paragraph("\n")
        doc.add_paragraph(f"🏆 {posisi}").runs[0].bold = True
        df_pos = ranking_df[ranking_df["Posisi"] == posisi].reset_index(drop=True)
        table = doc.add_table(rows=1, cols=4)
        hdr = table.rows[0].cells
        hdr[0].text = "No"
        hdr[1].text = "Nama"
        hdr[2].text = "Total Skor"
        hdr[3].text = "Penghargaan"

        for i, row in df_pos.iterrows():
            cells = table.add_row().cells
            cells[0].text = str(i + 1)
            cells[1].text = row["Nama"]
            cells[2].text = str(round(row["Total"], 2))
            if i == 0:
                cells[3].text = "🥇 Juara 1 (Emas)"
            elif i == 1:
                cells[3].text = "🥈 Juara 2 (Perak)"
            elif i == 2:
                cells[3].text = "🥉 Juara 3 (Perunggu)"
            else:
                cells[3].text = "-"

        doc.add_paragraph("\nUcapan Selamat:", style="Normal").runs[0].bold = True
        doc.add_paragraph(
            f"🎉 Selamat kepada {df_pos.iloc[0]['Nama']} atas pencapaian nilai tertinggi dan ditetapkan sebagai calon terbaik {posisi}.")

    doc.add_paragraph("\n\n")
    doc.add_paragraph("Lembar Pengesahan Penilai:").runs[0].bold = True
    pengesahan = doc.add_table(rows=1, cols=2)
    pengesahan.rows[0].cells[0].text = "Nama Penilai"
    pengesahan.rows[0].cells[1].text = "Tanda Tangan"

    for _, row in penilai_df.iterrows():
        r = pengesahan.add_row().cells
        r[0].text = row["Nama Penilai"]
        r[1].text = ".........................."

    qr = qrcode.make("Dokumen sah, diterbitkan oleh Panitia Pemilihan BUMDes Desa Keling")
    buffer = BytesIO()
    qr.save(buffer)
    buffer.seek(0)
    doc.add_paragraph("\n")
    doc.add_picture(buffer, width=Inches(1.5))
    doc.add_paragraph("Barcode ini menunjukkan dokumen resmi yang diterbitkan oleh Panitia Pemilihan Pengurus BUMDes Buwana Raharja Desa Keling.")

    path = f"{DATA_FOLDER}/Rekap_Final_Penilaian_BUMDes.docx"
    doc.save(path)

    with open(path, "rb") as f:
        st.download_button("📄 Download Word Final", f, file_name="Rekap_Final_Penilaian_BUMDes.docx")

st.markdown("<div style='text-align:center'>Developed by CV Mitra Utama Consultindo</div>", unsafe_allow_html=True)
