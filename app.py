import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Path data
DATA_FOLDER = "data"
PENILAI_FILE = os.path.join(DATA_FOLDER, "penilai.csv")
KANDIDAT_FILE = os.path.join(DATA_FOLDER, "kandidat.csv")
HASIL_FILE = os.path.join(DATA_FOLDER, "hasil_penilaian.csv")

# Bobot penilaian
bobot = {
    "Tes Psikologi": 0.15,
    "Tes MS Office": 0.15,
    "Presentasi Gagasan": 0.30,
    "Esai Refleksi Diri": 0.20,
    "Wawancara Panel": 0.20
}

st.set_page_config(page_title="Polling Pengurus BUMDes", layout="wide")
st.title("📊 Sistem Polling Pengurus BUMDes Buwana Raharja")

# --- Load Data ---
os.makedirs(DATA_FOLDER, exist_ok=True)

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

# --- Form Penilaian ---
st.subheader("📝 Form Penilaian")
with st.form("form_penilaian"):
    nama_penilai = st.selectbox("Pilih Identitas Penilai:", penilai_df["Nama Penilai"])
    posisi = st.selectbox("Pilih Posisi yang Dinilai:", kandidat_df["Posisi"].unique())
    kandidat_list = kandidat_df[kandidat_df["Posisi"] == posisi]["Nama"].tolist()
    kandidat_dipilih = st.selectbox("Pilih Kandidat:", kandidat_list)

    # Cek apakah sudah dinilai
    sudah_nilai = (
        (hasil_df["Penilai"] == nama_penilai)
        & (hasil_df["Posisi"] == posisi)
        & (hasil_df["Nama"] == kandidat_dipilih)
    ).any()

    nilai_input = {}
    if sudah_nilai:
        st.warning("⚠️ Penilaian untuk posisi dan kandidat ini oleh Anda sudah terkunci.")
    else:
        for aspek in bobot:
            nilai_input[aspek] = st.slider(f"{aspek} (0–100)", 0, 100, 0)

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
        st.success("✅ Penilaian berhasil disimpan dan terkunci.")

# --- Rekapitulasi ---
st.subheader("📈 Rekapitulasi Hasil Sementara")
if not hasil_df.empty:
    hasil_df["Total"] = hasil_df[[*bobot]].apply(
        lambda row: sum(row[aspek] * bobot[aspek] for aspek in bobot), axis=1)
    ranking_df = hasil_df.groupby(["Nama", "Posisi"]).agg({"Total": "mean"}).reset_index()
    ranking_df = ranking_df.sort_values(["Posisi", "Total"], ascending=[True, False])

    for posisi in ranking_df["Posisi"].unique():
        st.markdown(f"### 🏆 Hasil untuk {posisi}")
        st.dataframe(ranking_df[ranking_df["Posisi"] == posisi][["Nama", "Total"]].reset_index(drop=True))

    st.download_button(
        label="⬇️ Download Rekap Penilaian (.CSV)",
        data=ranking_df.to_csv(index=False).encode("utf-8"),
        file_name="rekap_penilaian.csv",
        mime="text/csv"
    )

# --- Footer ---
st.markdown("---")
st.markdown(
    "<div style='text-align: center;'>Developed by CV Mitra Utama Consultindo</div>",
    unsafe_allow_html=True
)
