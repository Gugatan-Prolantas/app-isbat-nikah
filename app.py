import io
import docx
import streamlit as st
from datetime import datetime, date

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
                    if placeholder in cell.text:
                        for p in cell.paragraphs:
                            if placeholder in p.text:
                                p.text = p.text.replace(placeholder, str_val)

def hitung_umur(born, target_date):
    """Menghitung umur berdasarkan tanggal lahir dan tanggal target."""
    if born is None or target_date is None:
        return 0
    return target_date.year - born.year - ((target_date.month, target_date.day) < (born.month, born.day))

def format_tanggal_indo(dt):
    """Mengubah format datetime.date ke teks Bahasa Indonesia (misal: 2 Agustus 2026)."""
    bulan_indo = [
        "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    return f"{dt.day} {bulan_indo[dt.month]} {dt.year}"

# Konfigurasi Tampilan Halaman
st.set_page_config(page_title="Form Permohonan Isbat Nikah", layout="wide")

st.title("📄 Generator Surat Permohonan Isbat Nikah")
st.write("Isi formulir di bawah ini untuk membuat dokumen permohonan Isbat Nikah secara otomatis.")

# Pilihan daftar pekerjaan baku
OPSI_PEKERJAAN = ["Pegawai BUMN/BUMD", "ASN", "Anggota Polri", "Anggota TNI", "Lain-lain"]

# Formulir Input
with st.form("form_isbat"):
    st.subheader("1. Informasi Permohonan")
    
    # Input tanggal permohonan menggunakan Date Picker (Default: Hari Ini)
    tgl_permohonan_val = st.date_input("Tanggal Surat Permohonan", value=date.today())
    tanggal_permohonan_str = format_tanggal_indo(tgl_permohonan_val)

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("2. Data Pemohon I (Suami)")
        nama_p1 = st.text_input("Nama Lengkap Pemohon I")
        nik_p1 = st.text_input("NIK Pemohon I")
        tempat_lahir_p1 = st.text_input("Tempat Lahir Pemohon I")
        
        # Date Picker untuk Tanggal Lahir Pemohon I
        tgl_lahir_p1_val = st.date_input(
            "Tanggal Lahir Pemohon I", 
            value=date(1990, 1, 1),
            min_value=date(1940, 1, 1),
            max_value=date.today()
        )
        tanggal_lahir_p1_str = format_tanggal_indo(tgl_lahir_p1_val)
        
        # Hitung Umur Otomatis
        umur_p1 = hitung_umur(tgl_lahir_p1_val, tgl_permohonan_val)
        st.info(f"💡 **Umur Pemohon I:** {umur_p1} tahun (dihitung otomatis)")

        # Pilihan Pekerjaan Pemohon I
        pilihan_pekerjaan_p1 = st.selectbox("Pekerjaan Pemohon I", OPSI_PEKERJAAN, key="pilihan_p1")
        if pilihan_pekerjaan_p1 == "Lain-lain":
            pekerjaan_p1 = st.text_input("Sebutkan Pekerjaan Pemohon I", value="", placeholder="Contoh: Wiraswasta / Petani / Ibu Rumah Tangga")
        else:
            pekerjaan_p1 = pilihan_pekerjaan_p1

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
        
        # Date Picker untuk Tanggal Lahir Pemohon II
        tgl_lahir_p2_val = st.date_input(
            "Tanggal Lahir Pemohon II", 
            value=date(1995, 1, 1),
            min_value=date(1940, 1, 1),
            max_value=date.today()
        )
        tanggal_lahir_p2_str = format_tanggal_indo(tgl_lahir_p2_val)
        
        # Hitung Umur Otomatis
        umur_p2 = hitung_umur(tgl_lahir_p2_val, tgl_permohonan_val)
        st.info(f"💡 **Umur Pemohon II:** {umur_p2} tahun (dihitung otomatis)")

        # Pilihan Pekerjaan Pemohon II
        pilihan_pekerjaan_p2 = st.selectbox("Pekerjaan Pemohon II", OPSI_PEKERJAAN, key="pilihan_p2")
        if pilihan_pekerjaan_p2 == "Lain-lain":
            pekerjaan_p2 = st.text_input("Sebutkan Pekerjaan Pemohon II", value="", placeholder="Contoh: Ibu Rumah Tangga / Wiraswasta")
        else:
            pekerjaan_p2 = pilihan_pekerjaan_p2

        pendidikan_p2 = st.text_input("Pendidikan Pemohon II")
        telepon_p2 = st.text_input("Nomor Telepon Pemohon II")
        email_p2 = st.text_input("Email Pemohon II")
        alamat_p2 = st.text_area("Alamat Lengkap Pemohon II")
        status_p2 = st.selectbox("Status Saat Menikah (Pemohon II)", ["Perawan", "Janda"])

    st.subheader("4. Detail Pernikahan Sirri")
    col3, col4 = st.columns(2)
    with col3:
        tgl_nikah_val = st.date_input("Tanggal Nikah Sirri", value=date(2015, 1, 10))
        tanggal_nikah_sirri_str = format_tanggal_indo(tgl_nikah_val)
        
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
        "tanggal_permohonan": tanggal_permohonan_str,
        "nama_pemohon1": nama_p1,
        "nik_pemohon1": nik_p1,
        "tempat_lahir_pemohon1": tempat_lahir_p1,
        "tanggal_lahir_pemohon1": tanggal_lahir_p1_str,
        "umur_pemohon1": umur_p1,
        "pekerjaan_pemohon1": pekerjaan_p1,
        "pendidikan_pemohon1": pendidikan_p1,
        "nomor_telepon_pemohon1": telepon_p1,
        "email_pemohon1": email_p1,
        "alamat_pemohon1": alamat_p1,
        "nama_pemohon2": nama_p2,
        "nik_pemohon2": nik_p2,
        "tempat_lahir_pemohon2": tempat_lahir_p2,
        "tanggal_lahir_pemohon2": tanggal_lahir_p2_str,
        "umur_pemohon2": umur_p2,
        "pekerjaan_pemohon2": pekerjaan_p2,
        "pendidikan_pemohon2": pendidikan_p2,
        "nomor_telepon_pemohon2": telepon_p2,
        "email_pemohon2": email_p2,
        "alamat_pemohon2": alamat_p2,
        "tanggal_nikah_sirri": tanggal_nikah_sirri_str,
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

        # Simpan hasil ke memori buffer
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
