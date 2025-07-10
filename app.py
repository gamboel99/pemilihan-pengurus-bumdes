import streamlit as st
import pandas as pd
import os
from datetime import datetime

# ---------- Konstanta & File Path ---------- #
DATA_FILE = "penilaian.csv"
KANDIDAT_FILE = "kandidat.csv"

# ---------- Inisialisasi Session State ---------- #
if "penilai" not in st.session_state:
    st.session_state.penilai = {}
if "reset_form" not in st.session_state:
    st.session_state.reset_form = False

# ---------- Judul ---------- #
st.title("📋 Polling Penilaian Pengurus BUMDes Buwana Raharja Desa Keling")
st.markdown("---")

# ---------- Input Identitas Penilai ---------- #
st.header("🧑‍⚖️ Identitas Penilai")
with st.form("form_identitas"):
    nama_penilai = st.text_input("Nama Penilai")
    jabatan = st.text_input("Jabatan")
    asal_lembaga = st.selectbox("Asal Lembaga", ["Pemdes", "BPD", "Kecamatan", "DPMPD"])
    simpan = st.form_submit_button("💾 Simpan Identitas")

if simpan:
    if nama_penilai == "" or jabatan == "":
        st.warning("Lengkapi semua kolom terlebih dahulu.")
    else:
        st.session_state.penilai = {
            "nama": nama_penilai,
            "jabatan": jabatan,
            "asal_lembaga": asal_lembaga
        }
        st.success("✅ Identitas disimpan. Scroll ke bawah untuk melanjutkan.")

# ---------- Lanjut Jika Identitas Sudah Diisi ---------- #
if st.session_state.penilai != {}:
    st.markdown("---")
    st.subheader("📝 Form Penilaian")

    # ---------- Load Data Kandidat ---------- #
    kandidat_df = pd.read_csv(KANDIDAT_FILE)
    posisi_list = kandidat_df["Posisi"].unique().tolist()

    posisi_dipilih = st.selectbox("Pilih Posisi", posisi_list)
    kandidat_list = kandidat_df[kandidat_df["Posisi"] == posisi_dipilih]["Nama"].tolist()
    kandidat_dipilih = st.selectbox("Pilih Nama Kandidat", kandidat_list)

    # Reset Form Saat Ganti Kandidat / Posisi
    if st.session_state.reset_form or "skor" not in st.session_state:
        st.session_state.skor = {"Psikologi": 0, "MS Office": 0, "Presentasi": 0, "Esai": 0, "Wawancara": 0}
        st.session_state.reset_form = False

    # ---------- Form Penilaian ---------- #
    with st.form("form_penilaian"):
        st.session_state.skor["Psikologi"] = st.number_input("Tes Psikologi (15%)", 0, 100, step=1, value=st.session_state.skor["Psikologi"])
        st.session_state.skor["MS Office"] = st.number_input("Kemampuan MS Office (15%)", 0, 100, step=1, value=st.session_state.skor["MS Office"])
        st.session_state.skor["Presentasi"] = st.number_input("Presentasi & Pemahaman Program (30%)", 0, 100, step=1, value=st.session_state.skor["Presentasi"])
        st.session_state.skor["Esai"] = st.number_input("Esai Gagasan & Refleksi Diri (20%)", 0, 100, step=1, value=st.session_state.skor["Esai"])
        st.session_state.skor["Wawancara"] = st.number_input("Wawancara Panel (20%)", 0, 100, step=1, value=st.session_state.skor["Wawancara"])

        submit = st.form_submit_button("✅ Simpan Penilaian")

    # ---------- Hitung dan Simpan ---------- #
    if submit:
        total = round(
            0.15 * st.session_state.skor["Psikologi"] +
            0.15 * st.session_state.skor["MS Office"] +
            0.30 * st.session_state.skor["Presentasi"] +
            0.20 * st.session_state.skor["Esai"] +
            0.20 * st.session_state.skor["Wawancara"], 2
        )

        new_row = pd.DataFrame([{**st.session_state.penilai,
                                  "Nama": kandidat_dipilih,
                                  "Posisi": posisi_dipilih,
                                  "Total": total,
                                  "Waktu": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}])

        # Cegah Penilaian Ganda
        if os.path.exists(DATA_FILE):
            hasil_df = pd.read_csv(DATA_FILE)
            cek = hasil_df[
                (hasil_df["Nama Penilai"] == st.session_state.penilai["nama"]) &
                (hasil_df["Posisi"] == posisi_dipilih) &
                (hasil_df["Nama"] == kandidat_dipilih)
            ]
            if not cek.empty:
                st.warning("⚠️ Anda sudah memberikan penilaian kepada kandidat ini di posisi tersebut.")
            else:
                hasil_df = pd.concat([hasil_df, new_row], ignore_index=True)
                hasil_df.to_csv(DATA_FILE, index=False)
                st.success("✅ Penilaian berhasil disimpan.")
                st.session_state.reset_form = True
                st.experimental_rerun()
        else:
            new_row.to_csv(DATA_FILE, index=False)
            st.success("✅ Penilaian berhasil disimpan.")
            st.session_state.reset_form = True
            st.experimental_rerun()

    # ---------- Tampilkan Rekap Sementara ---------- #
    if os.path.exists(DATA_FILE):
        hasil_df = pd.read_csv(DATA_FILE)
        st.markdown("---")
        st.subheader("📊 Rekap Penilaian Sementara")
        posisi_group = hasil_df.groupby(["Posisi", "Nama"]).agg({"Total": "mean"}).reset_index().sort_values(["Posisi", "Total"], ascending=[True, False])
        for posisi in posisi_group["Posisi"].unique():
            st.markdown(f"### 🧩 Posisi: {posisi}")
            df_p = posisi_group[posisi_group["Posisi"] == posisi].reset_index(drop=True)
            st.dataframe(df_p)

    # ---------- Footer ---------- #
    st.markdown("""
    <hr>
    <center>Developed by CV Mitra Utama Consultindo</center>
    """, unsafe_allow_html=True)
