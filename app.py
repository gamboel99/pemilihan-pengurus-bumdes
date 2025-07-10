# polling_bumdes_final.py
import streamlit as st
import pandas as pd
import os
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import qrcode
from io import BytesIO
from datetime import datetime

# ---------- Setup ----------
DATA_FOLDER = "data"
KANDIDAT_FILE = os.path.join(DATA_FOLDER, "kandidat.csv")
HASIL_FILE = os.path.join(DATA_FOLDER, "hasil_penilaian.csv")
LOGO_BUMDES = os.path.join(DATA_FOLDER, "logo_bumdes.png")
LOGO_DESA = os.path.join(DATA_FOLDER, "logo_desa.png")

os.makedirs(DATA_FOLDER, exist_ok=True)

if not os.path.exists(KANDIDAT_FILE):
    pd.DataFrame(columns=["Nama", "Posisi"]).to_csv(KANDIDAT_FILE, index=False)
if not os.path.exists(HASIL_FILE):
    pd.DataFrame(columns=["Nama", "Posisi", "Nama Penilai", "Jabatan", "Lembaga", "Psikologi", "Office", "Presentasi", "Esai", "Wawancara", "Total"]).to_csv(HASIL_FILE, index=False)

# ---------- Fungsi ----------
def reset_nilai():
    st.session_state["psikologi"] = 0
    st.session_state["office"] = 0
    st.session_state["presentasi"] = 0
    st.session_state["esai"] = 0
    st.session_state["wawancara"] = 0

def hitung_total(p, o, pre, e, w):
    return round((p*0.15) + (o*0.15) + (pre*0.3) + (e*0.2) + (w*0.2), 2)

# ---------- UI ----------
st.title("🗳️ Polling Pemilihan Pengurus BUMDes Buwana Raharja")
st.subheader("📋 Identitas Penilai")

with st.form("form_identitas"):
    nama_penilai = st.text_input("Nama Penilai")
    jabatan = st.text_input("Jabatan")
    lembaga = st.selectbox("Asal Lembaga", ["Pemdes", "BPD", "Kecamatan", "DPMPD"])
    simpan_id = st.form_submit_button("Simpan Identitas")

if simpan_id:
    st.session_state["penilai"] = {"nama": nama_penilai, "jabatan": jabatan, "lembaga": lembaga}
    reset_nilai()
    st.success("Identitas disimpan. Scroll ke bawah untuk melanjutkan.")

if "penilai" in st.session_state:
    kandidat_df = pd.read_csv(KANDIDAT_FILE)
    hasil_df = pd.read_csv(HASIL_FILE)
    penilai = st.session_state["penilai"]

    st.markdown("---")
    st.subheader("🧾 Form Penilaian Kandidat")

    posisi_list = kandidat_df["Posisi"].unique().tolist()
    posisi = st.selectbox("Pilih Posisi", posisi_list, on_change=reset_nilai)
    kandidat_list = kandidat_df[kandidat_df["Posisi"] == posisi]["Nama"].tolist()
    nama_kandidat = st.selectbox("Pilih Nama Kandidat", kandidat_list, on_change=reset_nilai)

    sudah_dinilai = hasil_df[
        (hasil_df["Nama Penilai"] == penilai["nama"]) &
        (hasil_df["Nama"] == nama_kandidat) &
        (hasil_df["Posisi"] == posisi)
    ]

    if not sudah_dinilai.empty:
        st.warning("Anda sudah menilai kandidat ini untuk posisi ini.")
    else:
        psikologi = st.number_input("Tes Psikologi (0-100)", 0, 100, key="psikologi")
        office = st.number_input("Tes MS Office (0-100)", 0, 100, key="office")
        presentasi = st.number_input("Presentasi (0-100)", 0, 100, key="presentasi")
        esai = st.number_input("Esai Gagasan (0-100)", 0, 100, key="esai")
        wawancara = st.number_input("Wawancara (0-100)", 0, 100, key="wawancara")

        if st.button("💾 Simpan Penilaian"):
            total = hitung_total(psikologi, office, presentasi, esai, wawancara)
            new_row = pd.DataFrame.from_dict({
                "Nama": [nama_kandidat],
                "Posisi": [posisi],
                "Nama Penilai": [penilai["nama"]],
                "Jabatan": [penilai["jabatan"]],
                "Lembaga": [penilai["lembaga"]],
                "Psikologi": [psikologi],
                "Office": [office],
                "Presentasi": [presentasi],
                "Esai": [esai],
                "Wawancara": [wawancara],
                "Total": [total]
            })
            hasil_df = pd.concat([hasil_df, new_row], ignore_index=True)
            hasil_df.to_csv(HASIL_FILE, index=False)
            reset_nilai()
            st.success("Penilaian disimpan. Scroll untuk melihat rekap.")

    # ---------- Rekap Per Posisi ----------
    st.markdown("---")
    st.subheader("📊 Rekap Sementara Per Posisi")
    for pos in posisi_list:
        st.markdown(f"### Posisi: {pos}")
        data_pos = hasil_df[hasil_df["Posisi"] == pos]
        if not data_pos.empty:
            rekap = data_pos.groupby("Nama")["Total"].mean().reset_index()
            rekap_sorted = rekap.sort_values(by="Total", ascending=False)
            st.table(rekap_sorted)

    # ---------- Download Laporan ----------
    st.markdown("---")
    st.subheader("📥 Unduh Rekapitulasi Penilaian")
    if st.button("⬇️ Download Word Rekap"):
        # ...kode pembuatan Word akan disambung di file generate_word.py
        st.warning("Fitur Word akan disambung ke file generate_word.py")
