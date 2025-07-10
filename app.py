import streamlit as st
import pandas as pd
import os
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import qrcode
from io import BytesIO
from PIL import Image

# ===== Konstanta =====
DATA_FOLDER = "data"
PENILAI_FILE = os.path.join(DATA_FOLDER, "penilai.csv")
KANDIDAT_FILE = os.path.join(DATA_FOLDER, "kandidat.csv")
HASIL_FILE = os.path.join(DATA_FOLDER, "hasil_penilaian.csv")
os.makedirs(DATA_FOLDER, exist_ok=True)

bobot = {
    "Tes Psikologi": 0.15,
    "Tes MS Office": 0.15,
    "Presentasi Gagasan": 0.30,
    "Esai Refleksi Diri": 0.20,
    "Wawancara Panel": 0.20
}

# ===== Data Awal =====
if not os.path.exists(KANDIDAT_FILE):
    pd.DataFrame({
        "Nama": [
            "Eko Wahyu Diantoro", "Racal Pudjo Alberto",
            "Nanda Rafiatul", "Wahida Nayla", "Fenny Alvionita",
            "Nani NurMahmudah", "Novia Lestari",
            "Jelvi Tri K", "Elton Priloro", "Azuma Zundana"
        ],
        "Posisi": [
            "Direktur Utama", "Direktur Utama",
            "Sekretaris", "Sekretaris", "Sekretaris",
            "Sekretaris", "Sekretaris",
            "Bendahara", "Bendahara", "Bendahara"
        ]
    }).to_csv(KANDIDAT_FILE, index=False)

if not os.path.exists(HASIL_FILE):
    pd.DataFrame(columns=["Penilai", "Posisi", "Nama"] + list(bobot.keys()) + ["Timestamp"]).to_csv(HASIL_FILE, index=False)

# ===== Fungsi Export Word =====
def generate_word_doc(ranking_df, identitas_penilai):
    doc = Document()

    # Kop
    header = doc.add_table(rows=1, cols=3)
    row = header.rows[0].cells
    try:
        row[0].paragraphs[0].add_run().add_picture("assets/logo_pemdes.png", width=Inches(1))
    except:
        row[0].text = "Logo Pemdes"
    row[1].text = "PEMERINTAH DESA KELING\nBUMDes BUWANA RAHARJA\nKecamatan Kepung, Kabupaten Kediri"
    row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    try:
        row[2].paragraphs[0].add_run().add_picture("assets/logo_bumdes.png", width=Inches(1))
    except:
        row[2].text = "Logo BUMDes"

    doc.add_paragraph("\nREKAPITULASI HASIL PENILAIAN CALON PENGURUS BUMDES", style='Heading 1').alignment = WD_ALIGN_PARAGRAPH.CENTER

    posisi_list = ranking_df["Posisi"].unique()
    for posisi in posisi_list:
        doc.add_paragraph(f"\nPosisi: {posisi}", style='Heading 2')
        df_posisi = ranking_df[ranking_df["Posisi"] == posisi].sort_values("Total", ascending=False)

        table = doc.add_table(rows=1, cols=3)
        table.style = 'Table Grid'
        table.columns[0].width = Inches(3)
        hdr = table.rows[0].cells
        hdr[0].text = "Nama"
        hdr[1].text = "Posisi"
        hdr[2].text = "Total"

        for idx, row in df_posisi.iterrows():
            cells = table.add_row().cells
            cells[0].text = row["Nama"]
            cells[1].text = row["Posisi"]
            cells[2].text = f"{row['Total']:.2f}"

        pemenang = df_posisi.iloc[0]["Nama"]
        doc.add_paragraph(f"\n\U0001F389 Selamat kepada {pemenang} terpilih sebagai {posisi} dengan skor tertinggi.", style="Intense Quote")

    # TTD dan Barcode
    doc.add_paragraph("\n\n")
    ttd = doc.add_table(rows=2, cols=2)
    ttd.cell(0,0).text = "Mengetahui,\nPanitia"
    ttd.cell(0,1).text = f"Kediri, {datetime.now().strftime('%d %B %Y')}\nPenilai"
    ttd.cell(1,0).text = "(...........................)"
    ttd.cell(1,1).text = f"({identitas_penilai['nama']})\n{identitas_penilai['jabatan']} - {identitas_penilai['instansi']}"

    doc.add_paragraph("\nDokumen resmi diterbitkan oleh Panitia Pemilihan Pengurus BUMDes Desa Keling.")
    qr = qrcode.make("Dokumen Resmi Panitia Pemilihan BUMDes Desa Keling")
    qr_io = BytesIO()
    qr.save(qr_io)
    qr_io.seek(0)
    doc.add_picture(qr_io, width=Inches(1))

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ===== Streamlit App =====
st.set_page_config("Polling Pengurus BUMDes", layout="wide")
st.title("📊 Polling Pengurus BUMDes Buwana Raharja")

# Identitas
st.header("🧾 Identitas Penilai")
with st.form("identitas"):
    nama = st.text_input("Nama Penilai")
    jabatan = st.text_input("Jabatan")
    instansi = st.selectbox("Asal Lembaga", ["Pemdes", "BPD", "Kecamatan", "DPMPD"])
    simpan = st.form_submit_button("Simpan Identitas")
    if simpan:
        st.session_state["penilai"] = {"nama": nama, "jabatan": jabatan, "instansi": instansi}
        st.success("Identitas disimpan. Scroll ke bawah untuk melanjutkan.")

# Penilaian
if "penilai" in st.session_state:
    kandidat_df = pd.read_csv(KANDIDAT_FILE)
    hasil_df = pd.read_csv(HASIL_FILE)

    st.header("📝 Penilaian Kandidat")
    posisi = st.selectbox("Pilih Posisi", kandidat_df["Posisi"].unique())
    kandidat_list = kandidat_df[kandidat_df["Posisi"] == posisi]["Nama"].tolist()
    kandidat = st.selectbox("Pilih Kandidat", kandidat_list)

    sudah_nilai = ((hasil_df["Penilai"] == st.session_state["penilai"]["nama"]) &
                   (hasil_df["Posisi"] == posisi) &
                   (hasil_df["Nama"] == kandidat)).any()

    with st.form("penilaian"):
        if sudah_nilai:
            st.warning("Anda sudah menilai kandidat ini.")
        else:
            nilai = {a: st.slider(a, 0, 100, key=f"{a}_{kandidat}") for a in bobot}
            simpan = st.form_submit_button("💾 Simpan Penilaian")
            if simpan:
                new = {
                    "Penilai": st.session_state["penilai"]["nama"],
                    "Posisi": posisi,
                    "Nama": kandidat,
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                new.update(nilai)
                hasil_df = pd.concat([hasil_df, pd.DataFrame([new])], ignore_index=True)
                hasil_df.to_csv(HASIL_FILE, index=False)
                st.success("✅ Penilaian tersimpan.")

    st.divider()
    st.header("📈 Rekapitulasi Hasil")
    if not hasil_df.empty:
        hasil_df["Total"] = hasil_df[[*bobot]].apply(lambda r: sum(r[a] * bobot[a] for a in bobot), axis=1)
        rekap = hasil_df.groupby(["Nama", "Posisi"]).agg({"Total": "mean"}).reset_index()
        rekap = rekap.sort_values(["Posisi", "Total"], ascending=[True, False])

        for pos in rekap["Posisi"].unique():
            st.subheader(f"Posisi: {pos}")
            df_pos = rekap[rekap["Posisi"] == pos][["Nama", "Total"]].reset_index(drop=True)
            st.dataframe(df_pos)

        st.download_button("⬇️ Download Rekap Word", data=generate_word_doc(rekap, st.session_state["penilai"]),
                           file_name="Rekap_Penilaian_BUMDes.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

st.markdown("---")
st.markdown("<div style='text-align:center;'>Developed by CV Mitra Utama Consultindo</div>", unsafe_allow_html=True)
