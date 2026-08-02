import io
import docx
import streamlit as st

def replace_text_in_document(doc, replacements):
    """Fungsi untuk mengganti teks placeholder di paragraf dan tabel."""
    for key, val in replacements.items():
        placeholder = f"[{key}]"
        str_val = str(val)
        
        # Ganti teks di paragraf utama
        for p in doc.paragraphs:
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, str_val)
                
        # Ganti teks jika ada dalam tabel
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if placeholder in p.text:
                            p.text = p.text.replace(placeholder, str_val)

# Konfigurasi Tampilan Halaman
st.set_page_config(page_title="Form Permohonan Isbat Nikah", layout="wide")

st.title("📄 Generator Surat Permohonan Isbat Nikah")
st.write("Isi formulir di bawah ini untuk membuat dokumen permohonan Isbat Nikah secara otomatis.")

# Formulir Input
with st.form("form_isbat"):
    st.subheader("1. Informasi Permohonan")
    tanggal_permohonan = st.text_input("Tanggal Surat Permohonan", value="2 Agustus 2026")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("2. Data Pemohon I (Suami)")
        nama_p1 = st.text_input("Nama Lengkap Pemohon I")
        nik_p1 = st.text_input("NIK Pemohon I")
        tempat_lahir_p1 = st.text_input("Tempat Lahir Pemohon I")
        tanggal_lahir_p1 = st.text_input("Tanggal Lahir Pemohon I (contoh: 12 Mei 1988)")
        umur_p1 = st.text_input("Umur Pemohon I (tahun)")
        pekerjaan_p1 = st.text_input("Pekerjaan Pemohon I")
        pendidikan_p1 = st.text_input("Pendidikan Pemohon I")
        telepon_p1 = st.text_input("Nomor Telepon Pemohon I")
        email_p1 = st.text_input("Email Pemohon I")
        alamat_p1 = st.text_area("Alamat Lengkap Pemohon I")
        status_p1 = st.selectbox("Status Saat Menikah (Pemohon I)", ["Jejaka", "Duda"])

    with col2:
        st.subheader("3. Data Pemohon II (Istri)")
        nama_p2 = st.text_input("Nama Lengkap Pemohon II")
        nik_p2 = st.text_input("NIK Pemohon II")
        tempat_lahir_p2 = st.text_input("Tempat Lahir Pemohon II")
        tanggal_lahir_p2 = st.text_input("Tanggal Lahir Pemohon II (contoh: 20 Agustus 1992)")
        umur_p2 = st.text_input("Umur Pemohon II (tahun)")
        pekerjaan_p2 = st.text_input("Pekerjaan Pemohon II")
        pendidikan_p2 = st.text_input("Pendidikan Pemohon II")
        telepon_p2 = st.text_input("Nomor Telepon Pemohon II")
        email_p2 = st.text_input("Email Pemohon II")
        alamat_p2 = st.text_area("Alamat Lengkap Pemohon II")
        status_p2 = st.selectbox("Status Saat Menikah (Pemohon II)", ["Perawan", "Janda"])

    st.subheader("4. Detail Pernikahan Sirri")
    col3, col4 = st.columns(2)
    with col3:
        tanggal_nikah_sirri = st.text_input("Tanggal Nikah Sirri", value="10 Januari 2015")
        tempat_nikah_sirri = st.text_input("Tempat Nikah Sirri", value="Purwokerto")
        hubungan_wali_nikah_p2 = st.text_input("Hubungan Wali Nikah dengan Pemohon II", value="Ayah Kandung")
        nama_wali = st.text_input("Nama Wali Nikah")
        alasan_wali_nikah = st.text_input("Keterangan Tambahan Wali Nikah (Opsional)", value="")

    with col4:
        yang_menikahkan = st.text_input("Nama Tokoh/Kiai yang Menikahkan")
        mahar = st.text_input("Mas Masuk / Mahar", value="Uang tunai Rp 500.000,- dan seperangkat alat shalat")
        saksi_nikah1 = st.text_input("Nama Saksi Nikah I")
        saksi_nikah2 = st.text_input("Nama Saksi Nikah II")

    st.subheader("5. Alasan & Tujuan Permohonan")
    jumlah_anak = st.text_input(
        "Keterangan Anak", 
        value="telah dikaruniai 2 orang anak masing-masing bernama ..."
    )
    alasan_tidak_mencatatkan_nikah = st.text_area(
        "Alasan Tidak Mencatatkan Nikah", 
        value="Bahwa Pemohon I dan Pemohon II tidak mendaftarkan pernikahannya ke KUA karena pertimbangan keterbatasan biaya pada saat itu."
    )
    alasan_mohon_isbat = st.text_input(
        "Maksud Permohonan Isbat Nikah", 
        value="penerbitan Buku Nikah serta pengurusan administrasi kependudukan anak"
    )

    submit_button = st.form_submit_button("🔨 Buat Dokumen")

# Proses Pembuatan Dokumen
if submit_button:
    # Mengumpulkan semua input ke dalam dictionary
    data_input = {
        "tanggal_permohonan": tanggal_permohonan,
        "nama_pemohon1": nama_p1,
        "nik_pemohon1": nik_p1,
        "tempat_lahir_pemohon1": tempat_lahir_p1,
        "tanggal_lahir_pemohon1": tanggal_lahir_p1,
        "umur_pemohon1": umur_p1,
        "pekerjaan_pemohon1": pekerjaan_p1,
        "pendidikan_pemohon1": pendidikan_p1,
        "nomor_telepon_pemohon1": telepon_p1,
        "email_pemohon1": email_p1,
        "alamat_pemohon1": alamat_p1,
        "nama_pemohon2": nama_p2,
        "nik_pemohon2": nik_p2,
        "tempat_lahir_pemohon2": tempat_lahir_p2,
        "tanggal_lahir_pemohon2": tanggal_lahir_p2,
        "umur_pemohon2": umur_p2,
        "pekerjaan_pemohon2": pekerjaan_p2,
        "pendidikan_pemohon2": pendidikan_p2,
        "nomor_telepon_pemohon2": telepon_p2,
        "email_pemohon2": email_p2,
        "alamat_pemohon2": alamat_p2,
        "tanggal_nikah_sirri": tanggal_nikah_sirri,
        "tempat_nikah_sirri": tempat_nikah_sirri,
        "hubungan_wali_nikah_p2": hubungan_wali_nikah_p2,
        "nama_wali": nama_wali,
        "alasan_wali_nikah": alasan_wali_nikah,
        "yang_menikahkan": yang_menikahkan,
        "mahar": mahar,
        "saksi_nikah1": saksi_nikah1,
        "saksi_nikah2": saksi_nikah2,
        "status_pemohon1": status_p1,
        "status_pemohon2": status_p2,
        "jumlah_anak": jumlah_anak,
        "alasan_tidak_mencatatkan_nikah": alasan_tidak_mencatatkan_nikah,
        "alasan_mohon_isbat": alasan_mohon_isbat,
    }

    try:
        # Buka file template
        doc = docx.Document("template_permohonan.docx")
        
        # Jalankan fungsi penggantian kata
        replace_text_in_document(doc, data_input)

        # Simpan hasil ke memori buffer untuk siap diunduh
        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)

        st.success("✅ Dokumen berhasil dibuat!")
        
        # Tombol Download
        file_name = f"Permohonan_Isbat_Nikah_{nama_p1 if nama_p1 else 'Draft'}.docx"
        st.download_button(
            label="📥 Download Surat Permohonan (.docx)",
            data=doc_io,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except FileNotFoundError:
        st.error("⚠️ File 'template_permohonan.docx' tidak ditemukan. Pastikan file template berada di folder yang sama dengan app.py.")
