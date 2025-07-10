# Di luar form: pilih penilai & posisi secara interaktif
st.subheader("📝 Form Penilaian")
nama_penilai = st.selectbox("Pilih Identitas Penilai:", penilai_df["Nama Penilai"])
posisi = st.selectbox("Pilih Posisi yang Dinilai:", kandidat_df["Posisi"].unique())
kandidat_list = kandidat_df[kandidat_df["Posisi"] == posisi]["Nama"].tolist()

with st.form("form_penilaian"):
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
            nilai_input[aspek] = st.number_input(
                f"{aspek} (0–100)", min_value=0, max_value=100, step=1, key=f"{aspek}_{posisi}_{kandidat_dipilih}"
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
        st.success("✅ Penilaian berhasil disimpan dan terkunci.")
