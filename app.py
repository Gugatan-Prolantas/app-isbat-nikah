import streamlit as st
import datetime
from io import BytesIO

# Try importing python-docx for Word file generation
try:
    import docx
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

st.set_page_config(
    page_title="Generator Surat Permohonan Isbat Nikah",
    page_icon="📄",
    layout="wide"
)

INDONESIAN_MONTHS = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]

def format_indo_date(dt):
    """Format datetime object into Indonesian date string."""
    if not dt:
        return ""
    return f"{dt.day} {INDONESIAN_MONTHS[dt.month - 1]} {dt.year}"

def calculate_age(birth_date, target_date):
    """Calculate age in years based on birth date and target date."""
    if not birth_date or not target_date:
        return 0
    age = target_date.year - birth_date.year
    if (target_date.month, target_date.day) < (birth_date.month, birth_date.day):
        age -= 1
    return max(0, age)

def build_letter_text(data):
    """Build formatted plain text for letter preview and text download."""
    p1 = data['p1']
    p2 = data['p2']
    
    tgl_surat_str = format_indo_date(data['tgl_permohonan'])
    
    # Children text
    anak_str = "belum / tidak dikaruniai anak."
    if data['status_anak'] == 'sudah' and data['anak_list']:
        anak_lines = [f"telah dikaruniai {len(data['anak_list'])} orang anak, yaitu:"]
        for idx, child in enumerate(data['anak_list'], 1):
            c_ttl = f"{child['tempat']}, {format_indo_date(child['tgl_lahir'])}"
            anak_lines.append(f"   {idx}. {child['nama']}, tempat/tgl lahir: {c_ttl} (umur {child['umur']} tahun)")
        anak_str = "\n".join(anak_lines)

    # Wali text
    wali_str = data['hubungan_wali']
    if data['hubungan_wali'] != 'Ayah Kandung' and data['alasan_wali']:
        wali_str += f" ({data['alasan_wali']})"

    # Handle optional status details gracefully to avoid blank lines
    status_detail_p1 = f"\n   Keterangan       : {p1['detail_status_text']}" if p1['detail_status_text'] else ""
    status_detail_p2 = f"\n   Keterangan       : {p2['detail_status_text']}" if p2['detail_status_text'] else ""

    text = f"""PERMOHONAN ISBAT NIKAH (PENGESAHAN NIKAH)

Hal : Permohonan Isbat Nikah
Purwokerto, {tgl_surat_str}

Kepada Yth.
Ketua Pengadilan Agama Purwokerto
di Tempat

Dengan hormat, kami yang bertanda tangan di bawah ini:

1. NAMA             : {p1['nama']}
   NIK              : {p1['nik']}
   Tempat/Tgl Lahir : {p1['tempat_lahir']}, {format_indo_date(p1['tgl_lahir'])} (Umur: {p1['umur']} tahun)
   Pekerjaan        : {p1['pekerjaan']}
   Pendidikan       : {p1['pendidikan']}
   Status           : {p1['status']}{status_detail_p1}
   Alamat           : {p1['alamat']}
   No. Telepon      : {p1['telepon']}
   Email            : {p1['email']}

2. NAMA             : {p2['nama']}
   NIK              : {p2['nik']}
   Tempat/Tgl Lahir : {p2['tempat_lahir']}, {format_indo_date(p2['tgl_lahir'])} (Umur: {p2['umur']} tahun)
   Pekerjaan        : {p2['pekerjaan']}
   Pendidikan       : {p2['pendidikan']}
   Status           : {p2['status']}{status_detail_p2}
   Alamat           : {p2['alamat']}
   No. Telepon      : {p2['telepon']}
   Email            : {p2['email']}

Selanjutnya disebut sebagai Pemohon I dan Pemohon II.

Menyatakan bahwa Pemohon I dan Pemohon II telah melangsungkan pernikahan secara syariat Islam pada tanggal {format_indo_date(data['tgl_nikah'])} di {data['tempat_nikah']}.
Pernikahan tersebut dilaksanakan dengan wali nikah {data['nama_wali']} selaku {wali_str}, di hadapan {data['yang_menikahkan']}, disaksikan oleh dua orang saksi bernama {data['saksi1']} dan {data['saksi2']}, dengan mas kawin/mahar berupa {data['mahar']}.

Bahwa selama pernikahan tersebut, Pemohon I dan Pemohon II {anak_str}

Bahwa {data['alasan_tidak_mencatatkan']}

Maksud permohonan ini adalah untuk {data['alasan_mohon']}.

Hormat kami,


Pemohon I                                        Pemohon II


( {p1['nama']} )                                 ( {p2['nama']} )
"""
    return text

def generate_docx(data):
    """Generate Word Document (.docx) in memory."""
    if not DOCX_AVAILABLE:
        return None

    doc = docx.Document()
    
    # Page setup
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    p1 = data['p1']
    p2 = data['p2']
    tgl_surat_str = format_indo_date(data['tgl_permohonan'])

    # Title
    title = doc.add_paragraph()
    title_run = title.add_run("PERMOHONAN ISBAT NIKAH (PENGESAHAN NIKAH)")
    title_run.bold = True
    title_run.font.size = Pt(14)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Header Date
    p_header = doc.add_paragraph()
    p_header.add_run(f"Hal : Permohonan Isbat Nikah\t\t\t\tPurwokerto, {tgl_surat_str}")

    doc.add_paragraph("\nKepada Yth.\nKetua Pengadilan Agama Purwokerto\ndi Tempat\n")
    doc.add_paragraph("Dengan hormat, kami yang bertanda tangan di bawah ini:")

    # Function to add applicant profile
    def add_applicant_data(num, applicant):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.add_run(f"{num}. Nama: ").bold = True
        p.add_run(f"{applicant['nama']}\n")
        p.add_run(f"   NIK: {applicant['nik']}\n")
        p.add_run(f"   TTL: {applicant['tempat_lahir']}, {format_indo_date(applicant['tgl_lahir'])} ({applicant['umur']} tahun)\n")
        p.add_run(f"   Pekerjaan: {applicant['pekerjaan']}\n")
        p.add_run(f"   Pendidikan: {applicant['pendidikan']}\n")
        p.add_run(f"   Status: {applicant['status']}\n")
        if applicant['detail_status_text']:
            p.add_run(f"   Keterangan: {applicant['detail_status_text']}\n")
        p.add_run(f"   Alamat: {applicant['alamat']}\n")
        p.add_run(f"   No. Telp: {applicant['telepon']}\n")
        p.add_run(f"   Email: {applicant['email']}")

    add_applicant_data(1, p1)
    add_applicant_data(2, p2)

    doc.add_paragraph("\nSelanjutnya disebut sebagai Pemohon I dan Pemohon II.\n")

    # Marriage details
    wali_str = data['hubungan_wali']
    if data['hubungan_wali'] != 'Ayah Kandung' and data['alasan_wali']:
        wali_str += f" ({data['alasan_wali']})"

    doc.add_paragraph(
        f"Menyatakan bahwa Pemohon I dan Pemohon II telah melangsungkan pernikahan secara syariat Islam "
        f"pada tanggal {format_indo_date(data['tgl_nikah'])} di {data['tempat_nikah']}. "
        f"Pernikahan tersebut dilaksanakan dengan wali nikah {data['nama_wali']} selaku {wali_str}, "
        f"di hadapan {data['yang_menikahkan']}, disaksikan oleh dua orang saksi bernama {data['saksi1']} dan {data['saksi2']}, "
        f"dengan mas kawin/mahar berupa {data['mahar']}."
    )

    # Children
    if data['status_anak'] == 'sudah' and data['anak_list']:
        p_child = doc.add_paragraph(f"Bahwa selama pernikahan tersebut, Pemohon I dan Pemohon II telah dikaruniai {len(data['anak_list'])} orang anak, yaitu:")
        for idx, child in enumerate(data['anak_list'], 1):
            c_ttl = f"{child['tempat']}, {format_indo_date(child['tgl_lahir'])}"
            doc.add_paragraph(f"{idx}. {child['nama']} (TTL: {c_ttl}, Umur: {child['umur']} tahun)", style='List Bullet')
    else:
        doc.add_paragraph("Bahwa selama pernikahan tersebut, Pemohon I dan Pemohon II belum / tidak dikaruniai anak.")

    doc.add_paragraph(f"Bahwa {data['alasan_tidak_mencatatkan']}")
    doc.add_paragraph(f"Maksud permohonan ini adalah untuk {data['alasan_mohon']}.\n")

    # Signatures
    p_sig = doc.add_paragraph()
    p_sig.add_run("Pemohon I\t\t\t\t\t\tPemohon II\n\n\n\n")
    p_sig.add_run(f"( {p1['nama']} )\t\t\t\t\t( {p2['nama']} )").bold = True

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

st.title("📄 Generator Surat Permohonan Isbat Nikah")
st.caption("Aplikasi Pembuat Surat Permohonan Isbat Nikah Otomatis (Streamlit Version)")

if not DOCX_AVAILABLE:
    st.warning("⚠️ Module `python-docx` tidak terdeteksi. Fitur unduh dokumen akan dialihkan ke format Teks (.txt). Jalankan `pip install python-docx` untuk mengaktifkan ekspor Word.")

col_form, col_preview = st.columns([7, 5])

with col_form:
    st.header("Formulir Permohonan")
    
    # --- 1. Informasi Permohonan ---
    with st.expander("1. Informasi Permohonan", expanded=True):
        tgl_permohonan = st.date_input("Tanggal Surat Permohonan", value=datetime.date.today())

    # --- 2. Data Pemohon I (Suami) ---
    with st.expander("2. Data Pemohon I (Suami)", expanded=True):
        c1, c2 = st.columns(2)
        nama_p1 = c1.text_input("Nama Lengkap Pemohon I", value="Ahmad Bin Fulan")
        nik_p1 = c2.text_input("NIK Pemohon I", value="3302123456780001")
        
        c3, c4 = st.columns(2)
        tempat_lahir_p1 = c3.text_input("Tempat Lahir (Pemohon I)", value="Banyumas")
        tgl_lahir_p1 = c4.date_input(
            "Tanggal Lahir (Pemohon I)", 
            value=datetime.date(1990, 1, 1),
            min_value=datetime.date(1900, 1, 1),
            max_value=datetime.date.today()
        )
        
        umur_p1 = calculate_age(tgl_lahir_p1, tgl_permohonan)
        st.info(f"💡 **Umur Pemohon I:** {umur_p1} tahun (dihitung otomatis)")

        pilihan_job_p1 = st.selectbox(
            "Pekerjaan Pemohon I", 
            ["Pegawai BUMN/BUMD", "ASN", "Anggota Polri", "Anggota TNI", "Lain-lain"]
        )
        pekerjaan_p1 = pilihan_job_p1
        if pilihan_job_p1 == "Lain-lain":
            pekerjaan_p1 = st.text_input("Sebutkan Pekerjaan Pemohon I", value="Wiraswasta")

        pendidikan_p1 = st.selectbox("Pendidikan Pemohon I", ["Tidak Sekolah", "TK", "SD", "SLTP", "SLTA", "D1", "D2", "D3", "D4", "S1", "S2", "S3"], index=4)
        status_p1 = st.selectbox("Status Saat Nikah (Pemohon I)", ["Jejaka", "Duda Cerai", "Duda Mati"])

        detail_status_p1_text = ""
        if status_p1 == "Duda Cerai":
            st.markdown("##### Detail Akta Cerai Pemohon I")
            no_ac = st.text_input("No. Akta Cerai (P1)", value="1234/AC/2020/PA.Pwt")
            tgl_ac = st.date_input("Tanggal Akta Cerai (P1)", value=datetime.date(2020, 5, 10))
            pa_penerbit = st.text_input("PA Penerbit (P1)", value="PA Purwokerto")
            detail_status_p1_text = f"Akta Cerai No: {no_ac}, tgl {format_indo_date(tgl_ac)} diterbitkan {pa_penerbit}"
        elif status_p1 == "Duda Mati":
            st.markdown("##### Detail Kematian Istri Terdahulu")
            tgl_mati = st.date_input("Tanggal Kematian Istri (P1)", value=datetime.date(2019, 1, 10))
            no_surat_m = st.text_input("No. Surat Kematian (P1)", value="474.3/01/2019")
            tgl_surat_m = st.date_input("Tanggal Surat Kematian (P1)", value=datetime.date(2019, 1, 15))
            penerbit_m = st.text_input("Penerbit Surat Kematian (P1)", value="Kepala Desa Purwokerto")
            detail_status_p1_text = f"Surat Kematian No: {no_surat_m} tgl {format_indo_date(tgl_surat_m)} (meninggal tgl {format_indo_date(tgl_mati)}) diterbitkan {penerbit_m}"

        telepon_p1 = st.text_input("No. Telepon/HP Pemohon I", value="08123456789")
        email_p1 = st.text_input("Email Pemohon I", value="ahmad@email.com")
        alamat_p1 = st.text_area("Alamat Lengkap Pemohon I", value="RT 01 RW 02, Desa Purwokerto, Kec. Purwokerto Barat, Kab. Banyumas")

    # --- 3. Data Pemohon II (Istri) ---
    with st.expander("3. Data Pemohon II (Istri)", expanded=True):
        c1, c2 = st.columns(2)
        nama_p2 = c1.text_input("Nama Lengkap Pemohon II", value="Siti Bintan")
        nik_p2 = c2.text_input("NIK Pemohon II", value="3302123456780002")
        
        c3, c4 = st.columns(2)
        tempat_lahir_p2 = c3.text_input("Tempat Lahir (Pemohon II)", value="Banyumas")
        tgl_lahir_p2 = c4.date_input(
            "Tanggal Lahir (Pemohon II)", 
            value=datetime.date(1995, 1, 1),
            min_value=datetime.date(1900, 1, 1),
            max_value=datetime.date.today()
        )
        
        umur_p2 = calculate_age(tgl_lahir_p2, tgl_permohonan)
        st.info(f"💡 **Umur Pemohon II:** {umur_p2} tahun (dihitung otomatis)")

        pilihan_job_p2 = st.selectbox(
            "Pekerjaan Pemohon II", 
            ["Pegawai BUMN/BUMD", "ASN", "Anggota Polri", "Anggota TNI", "Lain-lain"],
            index=4
        )
        pekerjaan_p2 = pilihan_job_p2
        if pilihan_job_p2 == "Lain-lain":
            pekerjaan_p2 = st.text_input("Sebutkan Pekerjaan Pemohon II", value="Ibu Rumah Tangga")

        pendidikan_p2 = st.selectbox("Pendidikan Pemohon II", ["Tidak Sekolah", "TK", "SD", "SLTP", "SLTA", "D1", "D2", "D3", "D4", "S1", "S2", "S3"], index=4)
        status_p2 = st.selectbox("Status Saat Nikah (Pemohon II)", ["Perawan", "Janda Cerai", "Janda Mati"])

        detail_status_p2_text = ""
        if status_p2 == "Janda Cerai":
            st.markdown("##### Detail Akta Cerai Pemohon II")
            no_ac2 = st.text_input("No. Akta Cerai (P2)", value="5678/AC/2021/PA.Pwt")
            tgl_ac2 = st.date_input("Tanggal Akta Cerai (P2)", value=datetime.date(2021, 8, 15))
            pa_penerbit2 = st.text_input("PA Penerbit (P2)", value="PA Purwokerto")
            detail_status_p2_text = f"Akta Cerai No: {no_ac2}, tgl {format_indo_date(tgl_ac2)} diterbitkan {pa_penerbit2}"
        elif status_p2 == "Janda Mati":
            st.markdown("##### Detail Kematian Suami Terdahulu")
            tgl_mati2 = st.date_input("Tanggal Kematian Suami (P2)", value=datetime.date(2020, 2, 10))
            no_surat_m2 = st.text_input("No. Surat Kematian (P2)", value="474.3/02/2020")
            tgl_surat_m2 = st.date_input("Tanggal Surat Kematian (P2)", value=datetime.date(2020, 2, 15))
            penerbit_m2 = st.text_input("Penerbit Surat Kematian (P2)", value="Kepala Desa Purwokerto")
            detail_status_p2_text = f"Surat Kematian No: {no_surat_m2} tgl {format_indo_date(tgl_surat_m2)} (meninggal tgl {format_indo_date(tgl_mati2)}) diterbitkan {penerbit_m2}"

        telepon_p2 = st.text_input("No. Telepon/HP Pemohon II", value="08987654321")
        email_p2 = st.text_input("Email Pemohon II", value="siti@email.com")
        alamat_p2 = st.text_area("Alamat Lengkap Pemohon II", value="RT 01 RW 02, Desa Purwokerto, Kec. Purwokerto Barat, Kab. Banyumas")

    # --- 4. Detail Pernikahan Sirri ---
    with st.expander("4. Detail Pernikahan Sirri", expanded=True):
        c1, c2 = st.columns(2)
        tgl_nikah = c1.date_input("Tanggal Nikah Sirri", value=datetime.date(2015, 1, 10))
        tempat_nikah = c2.text_input("Tempat Nikah Sirri", value="Purwokerto")

        hubungan_wali = st.selectbox(
            "Hubungan Wali Nikah",
            ["Ayah Kandung", "Saudara Kandung", "Kakek Kandung (Ayah dari Ayah)", "Paman Kandung (Saudara Ayah)", "Tidak Ada Hubungan Apapun/Hakim"]
        )
        alasan_wali = ""
        if hubungan_wali != "Ayah Kandung":
            alasan_wali = st.text_input("Alasan Wali Nikah (Bukan Ayah Kandung)", value="karena ayah kandung sudah meninggal dunia pada tahun 2010")

        nama_wali = st.text_input("Nama Wali Nikah", value="Bpk. Abdullah")
        yang_menikahkan = st.text_input("Nama Kiai/Tokoh yang Menikahkan", value="KH. Ahmad Sholeh")
        mahar = st.text_input("Mas Kawin / Mahar", value="Uang tunai Rp 500.000,- dan seperangkat alat shalat")
        
        c3, c4 = st.columns(2)
        saksi1 = c3.text_input("Saksi Nikah I", value="Bpk. Umar")
        saksi2 = c4.text_input("Saksi Nikah II", value="Bpk. Usman")

    # --- 5. Alasan & Tujuan Permohonan ---
    with st.expander("5. Alasan & Tujuan Permohonan", expanded=True):
        status_anak = st.selectbox("Keterangan Dikaruniai Anak", ["belum", "sudah"], index=1, format_func=lambda x: "Telah dikaruniai anak" if x == "sudah" else "Belum / Tidak dikaruniai anak")
        
        anak_list = []
        if status_anak == "sudah":
            jumlah_anak = st.number_input("Jumlah Anak", min_value=1, max_value=10, value=2)
            
            default_children_data = [
                {"nama": "Ahmad Raihan", "tempat": "Purwokerto", "tgl": datetime.date(2016, 5, 12)},
                {"nama": "Siti Aisyah", "tempat": "Purwokerto", "tgl": datetime.date(2019, 8, 20)}
            ]

            for i in range(int(jumlah_anak)):
                st.markdown(f"**Data Anak ke-{i+1}**")
                d_nama = default_children_data[i]["nama"] if i < len(default_children_data) else f"Anak Ke-{i+1}"
                d_tempat = default_children_data[i]["tempat"] if i < len(default_children_data) else "Purwokerto"
                d_tgl = default_children_data[i]["tgl"] if i < len(default_children_data) else datetime.date(2020, 1, 1)

                c1, c2, c3 = st.columns([4, 3, 3])
                c_nama = c1.text_input(f"Nama Anak #{i+1}", value=d_nama, key=f"c_nama_{i}")
                c_tempat = c2.text_input(f"Tempat Lahir #{i+1}", value=d_tempat, key=f"c_tempat_{i}")
                c_tgl = c3.date_input(
                    f"Tanggal Lahir #{i+1}", 
                    value=d_tgl, 
                    min_value=datetime.date(1950, 1, 1),
                    max_value=datetime.date.today(),
                    key=f"c_tgl_{i}"
                )
                
                c_umur = calculate_age(c_tgl, tgl_permohonan)
                st.caption(f"-> Umur Anak ke-{i+1}: **{c_umur} tahun**")

                anak_list.append({
                    "nama": c_nama,
                    "tempat": c_tempat,
                    "tgl_lahir": c_tgl,
                    "umur": c_umur
                })

        alasan_tidak_mencatatkan = st.selectbox(
            "Alasan Tidak Mencatatkan Nikah",
            [
                "Bahwa Pemohon I dan Pemohon II telah melaporkan pernikahannya kepada kayim untuk didaftarkan pada Kantor Urusan Agama, namun kayim tersebut tidak melanjutkan pendaftarannya ke Pembantu Pegawai Pencatat Nikah Kantor Urusan Agama;",
                "Bahwa Pemohon I dan Pemohon II telah melaporkan pernikahannya ke Pembantu Pegawai Pencatat Nikah setempat, namun Pembantu Pegawai Pencatat Nikah tersebut tidak melaporkan pencatatan pernikahan tersebut ke Kantor Urusan Agama;",
                "Bahwa Pemohon I dan Pemohon II tidak mendaftarkan pernikahannya ke KUA karena pertimbangan keterbatasan biaya pada saat itu;",
                "Bahwa Pemohon I dan Pemohon II pernah memiliki Buku Kutipan Akta Nikah namun hilang/rusak;"
            ],
            index=2
        )

        pil_alasan_mohon = st.selectbox(
            "Maksud Permohonan Isbat Nikah",
            [
                "penerbitan akta nikah Para Pemohon serta keperluan lainnya",
                "mengurus akta kelahiran anak Para Pemohon serta keperluan lainnya",
                "penerbitan Buku Nikah serta pengurusan administrasi kependudukan anak",
                "mendapatkan tunjangan pensiunan Veteran RI serta keperluan lainnya",
                "keperluan tersendiri"
            ],
            index=2
        )
        alasan_mohon = pil_alasan_mohon
        if pil_alasan_mohon == "keperluan tersendiri":
            alasan_mohon = st.text_input("Sebutkan Keperluan Tersendiri Anda", value="pendaftaran ibadah haji / waris")

form_data = {
    "tgl_permohonan": tgl_permohonan,
    "p1": {
        "nama": nama_p1,
        "nik": nik_p1,
        "tempat_lahir": tempat_lahir_p1,
        "tgl_lahir": tgl_lahir_p1,
        "umur": umur_p1,
        "pekerjaan": pekerjaan_p1,
        "pendidikan": pendidikan_p1,
        "status": status_p1,
        "detail_status_text": detail_status_p1_text,
        "telepon": telepon_p1,
        "email": email_p1,
        "alamat": alamat_p1
    },
    "p2": {
        "nama": nama_p2,
        "nik": nik_p2,
        "tempat_lahir": tempat_lahir_p2,
        "tgl_lahir": tgl_lahir_p2,
        "umur": umur_p2,
        "pekerjaan": pekerjaan_p2,
        "pendidikan": pendidikan_p2,
        "status": status_p2,
        "detail_status_text": detail_status_p2_text,
        "telepon": telepon_p2,
        "email": email_p2,
        "alamat": alamat_p2
    },
    "tgl_nikah": tgl_nikah,
    "tempat_nikah": tempat_nikah,
    "hubungan_wali": hubungan_wali,
    "alasan_wali": alasan_wali,
    "nama_wali": nama_wali,
    "yang_menikahkan": yang_menikahkan,
    "mahar": mahar,
    "saksi1": saksi1,
    "saksi2": saksi2,
    "status_anak": status_anak,
    "anak_list": anak_list,
    "alasan_tidak_mencatatkan": alasan_tidak_mencatatkan,
    "alasan_mohon": alasan_mohon
}

with col_preview:
    st.header("Pratinjau Dokumen Real-time")
    
    letter_text = build_letter_text(form_data)
    
    st.text_area("Hasil Draf Surat Permohonan", value=letter_text, height=620)

    st.subheader("Unduh Dokumen Hasil Generator")
    
    if DOCX_AVAILABLE:
        docx_file = generate_docx(form_data)
        st.download_button(
            label="📥 Unduh Dokumen Word (.docx)",
            data=docx_file,
            file_name=f"Permohonan_Isbat_Nikah_{nama_p1.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    
    st.download_button(
        label="📄 Unduh Teks Draf (.txt)",
        data=letter_text,
        file_name=f"Permohonan_Isbat_Nikah_{nama_p1.replace(' ', '_')}.txt",
        mime="text/plain",
        use_container_width=True
    )
