import streamlit as st
import pandas as pd
import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import qrcode
from io import BytesIO
from PIL import Image

# --- Konstanta ---
DATA_FOLDER = "data"
KANDIDAT_FILE = os.path.join(DATA_FOLDER, "kandidat.csv")
HASIL_FILE = os.path.join(DATA_FOLDER, "hasil_penilaian.csv")
LOGO_BUMDES = "logo_bumdes.png"
LOGO_DESA = "logo_desa.png"

os.makedirs(DATA_FOLDER, exist_ok=True)

# --- Load kandidat ---
if os.path.exists(KANDIDAT_FILE):
    kandidat_df = pd.read_csv(KANDIDAT_FILE)
else:
    kandidat_df = pd.DataFrame(columns=["Nama", "Posisi"])

# --- Identitas Penilai ---
st.title("🗳️ Sistem Penilaian Pengurus BUMDes Buwana Raharja")
st.markdown("---")

if "penilai_info" not in st.session_state:
    with st.form("form_penilai"):
        nama_penilai = st.text_input("Nama Penilai")
        jabatan_penilai = st.text_input("Jabatan")
        lembaga_penilai = st.selectbox("Asal Lembaga", ["Pemdes", "BPD", "Kecamatan", "DPMPD"])
        submit = st.form_submit_button("Simpan Identitas")
        if submit and nama_penilai and jabatan_penilai:
            st.session_state.penilai_info = {
                "nama": nama_penilai,
                "jabatan": jabatan_penilai,
                "lembaga": lembaga_penilai
            }
            st.success("✅ Identitas disimpan. Scroll ke bawah untuk melanjutkan.")
    st.stop()

penilai = st.session_state.penilai_info
st.info(f"Penilai: {penilai['nama']} ({penilai['jabatan']} - {penilai['lembaga']})")

# --- Load hasil ---
if os.path.exists(HASIL_FILE):
    hasil_df = pd.read_csv(HASIL_FILE)
else:
    hasil_df = pd.DataFrame()

# --- Kandidat yang belum dinilai ---
sudah_dinilai = hasil_df[hasil_df["Nama Penilai"] == penilai["nama"]][["Nama", "Posisi"]] if not hasil_df.empty else pd.DataFrame(columns=["Nama", "Posisi"])
kandidat_tersedia = pd.merge(kandidat_df, sudah_dinilai, on=["Nama", "Posisi"], how="left", indicator=True)
kandidat_tersedia = kandidat_tersedia[kandidat_tersedia["_merge"] == "left_only"].drop(columns="_merge")

if kandidat_tersedia.empty:
    st.success("✅ Anda telah menyelesaikan semua penilaian.")
else:
    st.markdown("---")
    st.subheader("📝 Form Penilaian")

    posisi_pilih = st.selectbox("Pilih Posisi", kandidat_tersedia["Posisi"].unique())
    kandidat_pilih = st.selectbox("Pilih Kandidat", kandidat_tersedia[kandidat_tersedia["Posisi"] == posisi_pilih]["Nama"].unique())

    with st.form("form_penilaian"):
        psikologi = st.number_input("Tes Psikologi (15%)", 0, 100, value=0, key="psikologi")
        office = st.number_input("Tes MS Office (15%)", 0, 100, value=0, key="office")
        presentasi = st.number_input("Presentasi Gagasan (30%)", 0, 100, value=0, key="presentasi")
        esai = st.number_input("Esai Refleksi Diri (20%)", 0, 100, value=0, key="esai")
        wawancara = st.number_input("Wawancara Panel (20%)", 0, 100, value=0, key="wawancara")
        simpan = st.form_submit_button("💾 Simpan Penilaian")

    if simpan:
        skor_total = (psikologi * 0.15 + office * 0.15 + presentasi * 0.3 + esai * 0.2 + wawancara * 0.2)
        data = {
            "Nama": kandidat_pilih,
            "Posisi": posisi_pilih,
            "Nama Penilai": penilai["nama"],
            "Jabatan": penilai["jabatan"],
            "Lembaga": penilai["lembaga"],
            "Tes Psikologi": psikologi,
            "Tes MS Office": office,
            "Presentasi Gagasan": presentasi,
            "Esai Refleksi Diri": esai,
            "Wawancara Panel": wawancara,
            "Total Skor": skor_total
        }
        hasil_df = pd.concat([hasil_df, pd.DataFrame([data])], ignore_index=True)
        hasil_df.to_csv(HASIL_FILE, index=False)

        # Reset nilai input manual
        for k in ["psikologi", "office", "presentasi", "esai", "wawancara"]:
            st.session_state.pop(k, None)

        st.success("✅ Penilaian disimpan. Silakan lanjut ke kandidat berikutnya.")
        st.info("Silakan pilih posisi dan kandidat lain untuk melanjutkan penilaian.")

    # Rekap penilaian oleh penilai
    if not hasil_df.empty:
        st.subheader("📊 Penilaian Anda")
        rekap_penilai = hasil_df[hasil_df["Nama Penilai"] == penilai["nama"]]
        for posisi in rekap_penilai["Posisi"].unique():
            st.markdown(f"**Posisi: {posisi}**")
            st.dataframe(rekap_penilai[rekap_penilai["Posisi"] == posisi][["Nama", "Total Skor"]])

# --- Tombol Download Rekap Final ---
st.markdown("---")
st.subheader("📥 Unduh Rekap Final Keseluruhan")

if st.button("📄 Download Word Rekapitulasi"):
    if hasil_df.empty:
        st.warning("Belum ada data penilaian.")
    else:
        doc = Document()
        section = doc.sections[0]
        section.left_margin = section.right_margin = Pt(36)

        header = doc.add_paragraph()
        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = header.add_run("LAPORAN REKAPITULASI PENILAIAN\nPENGURUS BUMDes BUWANA RAHARJA DESA KELING")
        run.bold = True
        run.font.size = Pt(14)

        for posisi in kandidat_df["Posisi"].unique():
            doc.add_paragraph(f"\nPosisi: {posisi}", style="List Bullet")
            df_posisi = hasil_df[hasil_df["Posisi"] == posisi]
            if df_posisi.empty:
                doc.add_paragraph("Belum ada penilaian.")
                continue
            grouped = df_posisi.groupby("Nama")["Total Skor"].mean().reset_index().sort_values(by="Total Skor", ascending=False)
            table = doc.add_table(rows=1, cols=3)
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'Peringkat'
            hdr_cells[1].text = 'Nama'
            hdr_cells[2].text = 'Skor Rata-rata'
            for i, row in enumerate(grouped.itertuples(), 1):
                cells = table.add_row().cells
                medal = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else str(i)
                cells[0].text = medal
                cells[1].text = row.Nama
                cells[2].text = f"{row._2:.2f}"
            pemenang = grouped.iloc[0]["Nama"]
            doc.add_paragraph(f"\nSelamat kepada {pemenang} yang terpilih sebagai {posisi}.")

        doc.add_paragraph("\n\nLembar Pengesahan:")
        table = doc.add_table(rows=1, cols=3)
        row = table.rows[0].cells
        row[0].text = f"Penilai:\n{penilai['nama']}\n{penilai['jabatan']}\n{penilai['lembaga']}"
        row[1].text = "Direktur BUMDes\n(...................)"
        row[2].text = "Kepala Desa\n(...................)"

        doc.add_paragraph("\n\nDokumen ini resmi diterbitkan oleh Panitia Pemilihan Pengurus BUMDes.")
        img = qrcode.make("Dokumen sah oleh Panitia Pemilihan BUMDes Desa Keling")
        buf = BytesIO()
        img.save(buf)
        doc.add_picture(BytesIO(buf.getvalue()), width=Pt(100))

        output = BytesIO()
        doc.save(output)
        st.download_button(
            label="📥 Klik untuk Unduh",
            data=output.getvalue(),
            file_name="Rekap_Penilaian_BUMDes.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

st.markdown("<center>Developed by CV Mitra Utama Consultindo</center>", unsafe_allow_html=True)
