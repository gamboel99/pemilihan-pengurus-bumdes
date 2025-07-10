# FINAL SCRIPT - app.py

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

# ============== SETUP =====================
st.set_page_config(page_title="Polling Pengurus BUMDes", layout="wide")
st.title("📊 Sistem Polling Pengurus BUMDes Buwana Raharja")

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

# ============== DUMMY DATA INIT (optional) =====================
if not os.path.exists(PENILAI_FILE):
    pd.DataFrame({"Nama Penilai": [
        "Kepala Desa", "Sekretaris Desa", "Ketua BPD", "Anggota BPD",
        "Kasi PMD Kecamatan", "Pendamping Kecamatan", "DPMPD"
    ]}).to_csv(PENILAI_FILE, index=False)

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

penilai_df = pd.read_csv(PENILAI_FILE)
kandidat_df = pd.read_csv(KANDIDAT_FILE)
hasil_df = pd.read_csv(HASIL_FILE)

# ============== FORM IDENTITAS =====================
st.subheader("🧾 Identitas Penilai")
with st.form("form_identitas"):
    nama_penilai_form = st.text_input("Nama Lengkap Penilai")
    jabatan_penilai = st.text_input("Jabatan")
    instansi_penilai = st.text_input("Instansi")
    submitted_id = st.form_submit_button("✔️ Simpan Identitas")
    if submitted_id:
        st.session_state["identitas_penilai"] = {
            "nama": nama_penilai_form,
            "jabatan": jabatan_penilai,
            "instansi": instansi_penilai
        }
        st.success("✅ Identitas disimpan.")

# ============== FORM PENILAIAN =====================
st.subheader("📝 Form Penilaian")

if "identitas_penilai" in st.session_state:
    nama_penilai = st.session_state["identitas_penilai"]["nama"]
    posisi = st.selectbox("Pilih Posisi yang Dinilai:", kandidat_df["Posisi"].unique())
    kandidat_list = kandidat_df[kandidat_df["Posisi"] == posisi]["Nama"].tolist()

    kandidat_dipilih = st.selectbox("Pilih Kandidat:", kandidat_list)

    if "last_selected" not in st.session_state:
        st.session_state["last_selected"] = {"posisi": None, "kandidat": None}

    # Reset nilai jika kandidat berubah
    if (st.session_state["last_selected"]["posisi"] != posisi or
        st.session_state["last_selected"]["kandidat"] != kandidat_dipilih):

        for aspek in bobot:
            key = f"{aspek}_{posisi}_{kandidat_dipilih}"
            if key in st.session_state:
                del st.session_state[key]

        st.session_state["last_selected"]["posisi"] = posisi
        st.session_state["last_selected"]["kandidat"] = kandidat_dipilih

    sudah_nilai = (
        (hasil_df["Penilai"] == nama_penilai) &
        (hasil_df["Posisi"] == posisi) &
        (hasil_df["Nama"] == kandidat_dipilih)
    ).any()

    with st.form("form_penilaian"):
        nilai_input = {}

        if sudah_nilai:
            st.warning("⚠️ Anda sudah menilai kandidat ini. Nilai tidak dapat diubah.")
        else:
            for aspek in bobot:
                nilai_input[aspek] = st.number_input(
                    f"{aspek} (0-100)", min_value=0, max_value=100, step=1,
                    key=f"{aspek}_{posisi}_{kandidat_dipilih}"
                )

        submitted = st.form_submit_button("💾 Simpan Penilaian")
        if submitted and not sudah_nilai:
            row = {
                "Penilai": nama_penilai,
                "Posisi": posisi,
                "Nama": kandidat_dipilih,
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            for aspek in bobot:
                row[aspek] = nilai_input[aspek]
            hasil_df = pd.concat([hasil_df, pd.DataFrame([row])], ignore_index=True)
            hasil_df.to_csv(HASIL_FILE, index=False)
            st.success("✅ Penilaian disimpan.")

# ============== REKAPITULASI =====================
st.subheader("📈 Rekapitulasi Hasil")

if not hasil_df.empty:
    hasil_df["Total"] = hasil_df[[*bobot]].apply(
        lambda row: sum(row[aspek] * bobot[aspek] for aspek in bobot), axis=1)

    ranking_df = hasil_df.groupby(["Nama", "Posisi"]).agg({"Total": "mean"}).reset_index()
    ranking_df = ranking_df.sort_values(["Posisi", "Total"], ascending=[True, False])

    for posisi in ranking_df["Posisi"].unique():
        st.markdown(f"### 🏆 {posisi}")
        st.dataframe(ranking_df[ranking_df["Posisi"] == posisi][["Nama", "Total"]])

# ============== FOOTER =====================
st.markdown("---")
st.markdown(
    "<div style='text-align: center;'>Developed by CV Mitra Utama Consultindo</div>",
    unsafe_allow_html=True
)
