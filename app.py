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
    
    # Anak text
    anak_str = "belum / tidak dikaruniai anak;"
    if data['status_anak'] == 'sudah' and data['anak_list']:
        anak_lines = [f"telah dikaruniai {len(data['anak_list'])} orang anak, yaitu:"]
        for idx, child in enumerate(data['anak_list'], 1):
            c_ttl = f"{child['tempat']}, {format_indo_date(child['tgl_lahir'])}"
            anak_lines.append(f"    {idx}. {child['nama']}, tempat/tgl lahir: {c_ttl} (umur {child['umur']} tahun)")
        anak_str = "\n".join(anak_lines) + ";"

    # Wali text
    alasan_wali_str = f" {data['alasan_wali']}" if data['hubungan_wali'] != 'Ayah Kandung' and data['alasan_wali'] else ""
    wali_str = f"{data['hubungan_wali']} Pemohon II bernama {data['nama_wali']}{alasan_wali_str}"

    # Status details text (if any)
    stat_p1 = f"{p1['status']} ({p1['detail_status_text']})" if p1['detail_status_text'] else p1['status']
    stat_p2 = f"{p2['status']} ({p2['detail_status_text']})" if p2['detail_status_text'] else p2['status']

    text = f"""Hal : Permohonan Isbat Nikah\t\tPurwokerto, {tgl_surat_str}

        Kepada
Yth. Ketua Pengadilan Agama Purwokerto
di
Purwokerto.

Assalamu Alaikum Wr. Wb.

Dengan hormat,
Yang bertanda tangan di bawah ini :

Nama\t\t\t: {p1['nama']} 
NIK\t\t\t: {p1['nik']}
Tempat Tgl Lahir\t: {p1['tempat_lahir']}, {format_indo_date(p1['tgl_lahir'])} (umur {p1['umur']} tahun)
Agama \t\t\t: Islam
Pekerjaan \t\t: {p1['pekerjaan']}
Pendidikan \t\t: {p1['pendidikan']}
Nomor telepon\t\t: {p1['telepon']}
Email\t\t\t: {p1['email']}
Alamat\t\t\t: {p1['alamat']}, selanjutnya disebut sebagai Pemohon I;

Nama\t\t\t: {p2['nama']} 
NIK\t\t\t: {p2['nik']}
Tempat Tgl Lahir\t: {p2['tempat_lahir']}, {format_indo_date(p2['tgl_lahir'])} (umur {p2['umur']} tahun)
Agama \t\t\t: Islam
Pekerjaan \t\t: {p2['pekerjaan']}
Pendidikan \t\t: {p2['pendidikan']}
Nomor telepon\t\t: {p2['telepon']}
Email\t\t\t: {p2['email']}
Alamat\t\t\t: {p2['alamat']}, selanjutnya disebut sebagai Pemohon II;

Dengan ini mengajukan pemohonan pengesahan nikah, dengan alasan sebagai berikut:

1.\tBahwa Pemohon I dan Pemohon II telah menikah menurut agama Islam pada tanggal {format_indo_date(data['tgl_nikah'])} di {data['tempat_nikah']} dengan wali nikah adalah {wali_str}, yang dinikahkan oleh {data['yang_menikahkan']}, dengan maskawin berupa {data['mahar']} yang dibayar tunai, dan dihadiri oleh dua orang saksi masing-masing bernama {data['saksi1']} dan {data['saksi2']};
2.\tBahwa saat menikah Pemohon I berstatus {stat_p1} dan Pemohon II berstatus {stat_p2};
3.\tBahwa antara Pemohon I dan Pemohon II tidak ada hubungan keluarga, baik sedarah maupun sesusuan serta Pemohon II juga tidak dalam pinangan laki-laki lain;
4.\tBahwa, pernikahan Pemohon I dengan Pemohon II telah dilaksanakan menurut hukum Islam serta tidak ada masyarakat yang menggugat atau yang meragukan keabsahan atau keberatan atas pernikahan Pemohon I dengan Pemohon II tersebut;
5.\tBahwa dari pernikahan tersebut, Pemohon I dan Pemohon II {anak_str}
6.\tBahwa antara Pemohon I dengan Pemohon II belum pernah terjadi perceraian;
7.\tBahwa, sampai sekarang Pemohon I dengan Pemohon II belum memiliki buku nikah, karena pernikahan Pemohon I dengan Pemohon II tidak terdaftar di Kantor Urusan Agama setempat;
8.\t{data['alasan_tidak_mencatatkan']}
9.\tBahwa, Pemohon I tidak mempunyai isteri yang lain, selain pemohon II;
10.\tBahwa, sekarang Pemohon I dan Pemohon II sangat membutuhkan bukti pernikahan tersebut, untuk mengurus {data['alasan_mohon']};
11.\tBahwa Pemohon I dan Pemohon II bersedia menanggung segala biaya yang ditimbulkan dari pengajuan perkara ini;

Bahwa berdasarkan alasan-alasan tersebut di atas para Pemohon mohon kepada Ketua Pengadilan Agama Purwokerto cq. Majelis hakim yang memeriksa perkara ini berkenan menetapkan sebagai berikut :

Primer :
1.\tMengabulkan permohonan para Pemohon;
2.\tMenyatakan sah perkawinan antara Pemohon I {p1['nama']} dengan Pemohon II, {p2['nama']} yang dilaksanakan pada tanggal {format_indo_date(data['tgl_nikah'])} di {data['tempat_nikah']};
3.\tMenetapkan biaya perkara menurut ketentuan hukum dan perundang-undangan yang berlaku;

Subsider :
-\tAtau bilamana majelis hakim yang memeriksa perkara ini berpendapat lain, mohon penetapan yang seadil-adilnya;

Demikian permohonan para Pemohon, dan atas terkabulnya para Pemohon ucapkan terima kasih.

Wassalam


Pemohon I, {p1['nama']} \t\t\t\t\tPemohon II, {p2['nama']}
"""
    return text

def generate_docx(data):
    """Generate Word Document (.docx) in memory."""
    if not DOCX_AVAILABLE:
        return None

    doc = docx.Document()
    
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)
    
    # Page setup
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    p1 = data['p1']
    p2 = data['p2']
    tgl_surat_str = format_indo_date(data['tgl_permohonan'])

    p_header = doc.add_paragraph()
    p_header.add_run(f"Hal : Permohonan Isbat Nikah\t\t\tPurwokerto, {tgl_surat_str}")

    doc.add_paragraph("\n\tKepada\nYth. Ketua Pengadilan Agama Purwokerto\ndi\nPurwokerto.\n")
    
    doc.add_paragraph("Assalamu Alaikum Wr. Wb.\n")
    doc.add_paragraph("Dengan hormat,\nYang bertanda tangan di bawah ini :")

    # Helper function to add biodata with tabs
    def add_biodata_docx(applicant, title):
        p = doc.add_paragraph()
        fmt = (
            f"Nama\t\t\t: {applicant['nama']}\n"
            f"NIK\t\t\t: {applicant['nik']}\n"
            f"Tempat Tgl Lahir\t: {applicant['tempat_lahir']}, {format_indo_date(applicant['tgl_lahir'])} (umur {applicant['umur']} tahun)\n"
            f"Agama\t\t\t: Islam\n"
            f"Pekerjaan\t\t: {applicant['pekerjaan']}\n"
            f"Pendidikan\t\t: {applicant['pendidikan']}\n"
            f"Nomor telepon\t: {applicant['telepon']}\n"
            f"Email\t\t\t: {applicant['email']}\n"
            f"Alamat\t\t\t: {applicant['alamat']}, selanjutnya disebut sebagai {title};\n"
        )
        p.add_run(fmt)

    add_biodata_docx(p1, "Pemohon I")
    add_biodata_docx(p2, "Pemohon II")

    doc.add_paragraph("Dengan ini mengajukan pemohonan pengesahan nikah, dengan alasan sebagai berikut:")

    # Status details
    stat_p1 = f"{p1['status']} ({p1['detail_status_text']})" if p1['detail_status_text'] else p1['status']
    stat_p2 = f"{p2['status']} ({p2['detail_status_text']})" if p2['detail_status_text'] else p2['status']

    # Wali text
    alasan_wali_str = f" {data['alasan_wali']}" if data['hubungan_wali'] != 'Ayah Kandung' and data['alasan_wali'] else ""
    wali_str = f"{data['hubungan_wali']} Pemohon II bernama {data['nama_wali']}{alasan_wali_str}"
    
    # Children text
    anak_str = "belum / tidak dikaruniai anak;"
    if data['status_anak'] == 'sudah' and data['anak_list']:
        anak_lines = [f"telah dikaruniai {len(data['anak_list'])} orang anak, yaitu:"]
        for idx, child in enumerate(data['anak_list'], 1):
            c_ttl = f"{child['tempat']}, {format_indo_date(child['tgl_lahir'])}"
            anak_lines.append(f"    {idx}. {child['nama']}, tempat/tgl lahir: {c_ttl} (umur {child['umur']} tahun)")
        anak_str = "\n".join(anak_lines) + ";"

    # Setup Posita points
    posita_points = [
        f"Bahwa Pemohon I dan Pemohon II telah menikah menurut agama Islam pada tanggal {format_indo_date(data['tgl_nikah'])} di {data['tempat_nikah']} dengan wali nikah adalah {wali_str}, yang dinikahkan oleh {data['yang_menikahkan']}, dengan maskawin berupa {data['mahar']} yang dibayar tunai, dan dihadiri oleh dua orang saksi masing-masing bernama {data['saksi1']} dan {data['saksi2']};",
        f"Bahwa saat menikah Pemohon I berstatus {stat_p1} dan Pemohon II berstatus {stat_p2};",
        "Bahwa antara Pemohon I dan Pemohon II tidak ada hubungan keluarga, baik sedarah maupun sesusuan serta Pemohon II juga tidak dalam pinangan laki-laki lain;",
        "Bahwa, pernikahan Pemohon I dengan Pemohon II telah dilaksanakan menurut hukum Islam serta tidak ada masyarakat yang menggugat atau yang meragukan keabsahan atau keberatan atas pernikahan Pemohon I dengan Pemohon II tersebut;",
        f"Bahwa dari pernikahan tersebut, Pemohon I dan Pemohon II {anak_str}",
        "Bahwa antara Pemohon I dengan Pemohon II belum pernah terjadi perceraian;",
        "Bahwa, sampai sekarang Pemohon I dengan Pemohon II belum memiliki buku nikah, karena pernikahan Pemohon I dengan Pemohon II tidak terdaftar di Kantor Urusan Agama setempat;",
        data['alasan_tidak_mencatatkan'],
        "Bahwa, Pemohon I tidak mempunyai isteri yang lain, selain pemohon II;",
        f"Bahwa, sekarang Pemohon I dan Pemohon II sangat membutuhkan bukti pernikahan tersebut, untuk mengurus {data['alasan_mohon']};",
        "Bahwa Pemohon I dan Pemohon II bersedia menanggung segala biaya yang ditimbulkan dari pengajuan perkara ini;"
    ]

    # Add Posita with Numbering and Justified Alignment
    for i, point_text in enumerate(posita_points, 1):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        # Hanging indent logic (0.3 inches indentation for the text, negative for the number)
        p.paragraph_format.left_indent = Inches(0.3)
        p.paragraph_format.first_line_indent = Inches(-0.3)
        p.add_run(f"{i}.\t{point_text}")

    # Add Petitum
    p_petitum_intro = doc.add_paragraph()
    p_petitum_intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_petitum_intro.add_run("Bahwa berdasarkan alasan-alasan tersebut di atas para Pemohon mohon kepada Ketua Pengadilan Agama Purwokerto cq. Majelis hakim yang memeriksa perkara ini berkenan menetapkan sebagai berikut :")
    
    p_petitum = doc.add_paragraph()
    p_petitum.add_run("Primer :\n")
    p_petitum.add_run("1.\tMengabulkan permohonan para Pemohon;\n")
    p_petitum.add_run(f"2.\tMenyatakan sah perkawinan antara Pemohon I {p1['nama']} dengan Pemohon II, {p2['nama']} yang dilaksanakan pada tanggal {format_indo_date(data['tgl_nikah'])} di {data['tempat_nikah']};\n")
    p_petitum.add_run("3.\tMenetapkan biaya perkara menurut ketentuan hukum dan perundang-undangan yang berlaku;\n\n")
    p_petitum.add_run("Subsider :\n")
    p_petitum.add_run("-\tAtau bilamana majelis hakim yang memeriksa perkara ini berpendapat lain, mohon penetapan yang seadil-adilnya;")

    doc.add_paragraph("Demikian permohonan para Pemohon, dan atas terkabulnya para Pemohon ucapkan terima kasih.\n\nWassalam\n\n")

    # Signatures Table for alignment
    table = doc.add_table(rows=3, cols=2)
    table.autofit = True
    
    cell_1 = table.cell(0, 0)
    cell_2 = table.cell(0, 1)
    
    cell_1.paragraphs[0].text = "Pemohon I,"
    cell_2.paragraphs[0].text = "Pemohon II,"
    
    # Add empty spacing for signature
    table.cell(1, 0).paragraphs[0].add_run("\n\n")
    
    cell_1_name = table.cell(2, 0)
    cell_2_name = table.cell(2, 1)
    
    cell_1_name.paragraphs[0].text = p1['nama']
    cell_2_name.paragraphs[0].text = p2['nama']

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
