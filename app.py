import streamlit as st
import pandas as pd
import os

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

st.set_page_config(page_title="Polling Pemilihan Pengurus BUMDes", layout="wide")
st.title("📊 Sistem Polling Pemilihan Pengurus BUMDes Buwana Raharja")

# --- Load Data ---
os.makedirs(DATA_FOLDER, exist_ok=True)

if not os.path.exists(PENILAI_FILE):
    pd.DataFrame({"Nama Penilai": [
        "Kepala Desa", "Sekretaris Desa", "Ketua BPD", "Anggota BPD",
        "Kasi PMD Kecamatan", "Pendamping Kecamatan", "DPMPD"
    ]}).to_csv(PENILAI_FILE, index=False)

if not os.path.exists(KANDIDAT_FILE):
    pd.DataFrame({
        "Nama": ["Ahmad", "Budi", "Citra"],
        "Posisi": ["Direktur Utama", "Sekretaris", "Bendahara"]
    }).to_csv(KANDIDAT_FILE, index=False)

if not os.path.exists(HASIL_FILE):
    pd.DataFrame(columns=["Penilai", "Nama", "Posisi"] + list(bobot.keys())).to_csv(HASIL_FILE, index=False)

penilai_df = pd.read_csv(PENILAI_FILE)
kandidat_df = pd.read_csv(KANDIDAT_FILE)
hasil_df = pd.read_csv(HASIL_FILE)

# --- Form Penilaian ---
st.subheader("📝 Form Penilaian")
with st.form("form_penilaian"):
    nama_penilai = st.selectbox("Pilih Identitas Penilai:", penilai_df["Nama Penilai"])
    posisi = st.selectbox("Pilih Posisi yang Dinilai:", kandidat_df["Posisi"].unique())
    kandidat_list = kandidat_df[kandidat_df["Posisi"] == posisi]["Nama"].tolist()

    nilai_input = {}
    for kandidat in kandidat_list:
        st.markdown(f"### Kandidat: **{kandidat}**")
        nilai_input[kandidat] = {}
        for aspek in bobot:
            key = f"{kandidat}_{aspek}"
            nilai_input[kandidat][aspek] = st.slider(f"{aspek} untuk {kandidat} (0-100)", 0, 100, 0, key=key)

    submitted = st.form_submit_button("💾 Simpan Penilaian")

    if submitted:
        for kandidat in kandidat_list:
            row = {
                "Penilai": nama_penilai,
                "Nama": kandidat,
                "Posisi": posisi
            }
            for aspek in bobot:
                row[aspek] = nilai_input[kandidat][aspek]
            hasil_df = pd.concat([hasil_df, pd.DataFrame([row])], ignore_index=True)
        hasil_df.to_csv(HASIL_FILE, index=False)
        st.success("Penilaian berhasil disimpan!")

# --- Rekapitulasi ---
st.subheader("📈 Hasil Rekapitulasi")
if not hasil_df.empty:
    hasil_df["Total"] = hasil_df[[*bobot]].apply(lambda row: sum(row[aspek] * bobot[aspek] for aspek in bobot), axis=1)
    ranking_df = hasil_df.groupby(["Nama", "Posisi"]).agg({"Total": "mean"}).reset_index()
    ranking_df = ranking_df.sort_values(["Posisi", "Total"], ascending=[True, False])

    for posisi in ranking_df["Posisi"].unique():
        st.markdown(f"### 🏆 Hasil untuk {posisi}")
        st.dataframe(ranking_df[ranking_df["Posisi"] == posisi][["Nama", "Total"]].reset_index(drop=True))

# --- Footer ---
st.markdown("---")
st.markdown(
    "<div style='text-align: center;'>Developed by CV Mitra Utama Consultindo</div>",
    unsafe_allow_html=True
)
