import streamlit as st
import pandas as pd
import os
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import qrcode
from io import BytesIO

# --- Konstanta
DATA_FOLDER = "data"
KANDIDAT_FILE = os.path.join(DATA_FOLDER, "kandidat.csv")
HASIL_FILE = os.path.join(DATA_FOLDER, "hasil_penilaian.csv")
os.makedirs(DATA_FOLDER, exist_ok=True)

# --- Load Data Kandidat
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
        st.rerun()
    st.stop()
else:
    penilai = st.session_state.penilai_info
    st.success(f"Penilai: {penilai['nama']} ({penilai['jabatan']} - {penilai['lembaga']})")

# --- Filter Kandidat yang Belum Dinilai
if os.path.exists(HASIL_FILE):
    hasil_df = pd.read_csv(HASIL_FILE)
    sudah_dinilai = hasil_df[hasil_df["Nama Penilai"] == penilai["nama"]][["Nama", "Posisi"]]
    kandidat_tersedia = pd.merge(kandidat_df, sudah_dinilai, on=["Nama", "Posisi"], how="left", indicator=True)
    kandidat_tersedia = kandidat_tersedia[kandidat_tersedia["_merge"] == "left_only"].drop(columns=["_merge"])
else:
    hasil_df = pd.DataFrame()
    kandidat_tersedia = kandidat_df.copy()

if kandidat_tersedia.empty:
    st.info("✅ Anda telah menilai semua kandidat. Terima kasih.")
    st.stop()

# --- Pilih Posisi dan Kandidat
st.subheader("📝 Form Penilaian")
posisi_pilih = st.selectbox("Pilih Posisi:", kandidat_tersedia["Posisi"].unique())
kandidat_pilih = st.selectbox("Pilih Kandidat:", kandidat_tersedia[kandidat_tersedia["Posisi"] == posisi_pilih]["Nama"].unique())

# --- Reset skor saat ganti kandidat
if "nilai_skor" not in st.session_state:
    st.session_state.nilai_skor = {}
if kandidat_pilih != st.session_state.nilai_skor.get("kandidat"):
    st.session_state.nilai_skor = {"kandidat": kandidat_pilih}

# --- Form Penilaian Skor
with st.form("form_penilaian"):
    psikologi = st.number_input("Tes Psikologi (0-100)", 0, 100, value=0)
    office = st.number_input("Tes Microsoft Office (0-100)", 0, 100, value=0)
    presentasi = st.number_input("Presentasi Gagasan (0-100)", 0, 100, value=0)
    esai = st.number_input("Esai Refleksi Diri (0-100)", 0, 100, value=0)
    wawancara = st.number_input("Wawancara Panel (0-100)", 0, 100, value=0)
    simpan = st.form_submit_button("💾 Simpan Penilaian")

if simpan:
    nilai = pd.DataFrame([{
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
    hasil_df = pd.concat([hasil_df, nilai], ignore_index=True)
    hasil_df.to_csv(HASIL_FILE, index=False)
    st.success("✅ Penilaian berhasil disimpan.")
    st.rerun()

# --- Export Word Rekap
st.subheader("📄 Export Rekap Penilaian (Word)")
if st.button("📥 Generate Rekap Word"):
    if not os.path.exists(HASIL_FILE):
        st.warning("Belum ada data penilaian.")
    else:
        df = pd.read_csv(HASIL_FILE)
        df["Total"] = df[["Tes Psikologi", "Tes MS Office", "Presentasi Gagasan", "Esai Refleksi Diri", "Wawancara Panel"]].apply(
            lambda r: r["Tes Psikologi"]*0.15 + r["Tes MS Office"]*0.15 + r["Presentasi Gagasan"]*0.3 + r["Esai Refleksi Diri"]*0.2 + r["Wawancara Panel"]*0.2,
            axis=1
        )
        rekap = df.groupby(["Nama", "Posisi"]).agg({"Total": "mean"}).reset_index()
        rekap = rekap.sort_values(["Posisi", "Total"], ascending=[True, False])

        doc = Document()
        title = doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run("LAPORAN HASIL PENILAIAN\nPENGURUS BUMDes Buwana Raharja Desa Keling")
        run.bold = True
        run.font.size = Pt(14)

        for posisi in rekap["Posisi"].unique():
            doc.add_paragraph("\n")
            doc.add_paragraph(posisi).runs[0].bold = True
            table = doc.add_table(rows=1, cols=4)
            hdr = table.rows[0].cells
            hdr[0].text = "No"
            hdr[1].text = "Nama"
            hdr[2].text = "Total Skor"
            hdr[3].text = "Penghargaan"
            data_posisi = rekap[rekap["Posisi"] == posisi].reset_index(drop=True)
            for i, row in data_posisi.iterrows():
                cells = table.add_row().cells
                cells[0].text = str(i+1)
                cells[1].text = row["Nama"]
                cells[2].text = f"{row['Total']:.2f}"
                cells[3].text = ["🥇 Juara 1", "🥈 Juara 2", "🥉 Juara 3"][i] if i < 3 else "-"
            doc.add_paragraph(f"🎉 Selamat kepada {data_posisi.iloc[0]['Nama']} sebagai {posisi} terbaik.")

        doc.add_paragraph("\n\nLembar Pengesahan Penilai:").runs[0].bold = True
        table = doc.add_table(rows=1, cols=3)
        table.rows[0].cells[0].text = "Nama Penilai"
        table.rows[0].cells[1].text = "Jabatan"
        table.rows[0].cells[2].text = "Tanda Tangan"
        for _, r in df[["Nama Penilai", "Jabatan"]].drop_duplicates().iterrows():
            row = table.add_row().cells
            row[0].text = r["Nama Penilai"]
            row[1].text = r["Jabatan"]
            row[2].text = ".............................."

        qr = qrcode.make("Dokumen sah - Pemilihan Pengurus BUMDes Desa Keling")
        buf = BytesIO()
        qr.save(buf)
        buf.seek(0)
        doc.add_picture(buf, width=Inches(1.3))
        doc.add_paragraph("Barcode ini menunjukkan dokumen resmi yang diterbitkan oleh Panitia Pemilihan Pengurus BUMDes Buwana Raharja Desa Keling.")

        doc_path = os.path.join(DATA_FOLDER, "Rekap_Final_Penilaian_BUMDes.docx")
        doc.save(doc_path)
        with open(doc_path, "rb") as f:
            st.download_button("📄 Download Word", f, file_name="Rekap_Penilaian_BUMDes.docx")

# --- Footer
st.markdown("<div style='text-align:center'>Developed by CV Mitra Utama Consultindo</div>", unsafe_allow_html=True)
