import streamlit as st
from streamlit_drawable_canvas import st_canvas
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from PIL import Image
import io
import datetime
import re
import pandas as pd

LIST_PANGKAT = [
    "Pembina Utama (IV/e)", "Pembina Utama Madya (IV/d)", "Pembina Utama Muda (IV/c)",
    "Pembina Tingkat I (IV/b)", "Pembina (IV/a)", "Penata Tingkat I (III/d)",
    "Penata (III/c)", "Penata Muda Tingkat I (III/b)", "Penata Muda (III/a)"
]

LIST_JABATAN = [
    "Guru Ahli Muda",
    "Guru Ahli Pertama",
    "Guru Ahli Madya",
    "Bendahara",
    "Analis Sumber Daya Manusia Aparatur Ahli Muda",
]

JABATAN_MAP = {
    "Guru Muda": "Guru Ahli Muda",
    "Guru Pertama": "Guru Ahli Pertama",
    "Guru Ahli Muda": "Guru Ahli Muda",
    "Guru Ahli Pertama": "Guru Ahli Pertama",
    "Guru Ahli Madya": "Guru Ahli Madya",
    "Bendahara": "Bendahara",
    "Analis Sumber Daya Manusia Aparatur Ahli Muda": "Analis Sumber Daya Manusia Aparatur Ahli Muda",
}

PANGKAT_MAP = {
    "Pembina Utama (IV/e)": "Pembina Utama (IV/e)",
    "Pembina Utama Madya (IV/d)": "Pembina Utama Madya (IV/d)",
    "Pembina Utama Muda (IV/c)": "Pembina Utama Muda (IV/c)",
    "Pembina TK. I IV/b": "Pembina Tingkat I (IV/b)",
    "Pembina IV/a": "Pembina (IV/a)",
    "Penata TK. I III/d": "Penata Tingkat I (III/d)",
    "IX": "Penata (III/c)",
    "Penata Muda TK. I III/b": "Penata Muda Tingkat I (III/b)",
    "Penata Muda (III/a)": "Penata Muda (III/a)",
}

def normalisasi_teks(value):
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\u00a0", " ").strip())

@st.cache_data(show_spinner=False)
def muat_data_excel():
    try:
        df = pd.read_excel("data_pegawai.xlsx")
        required_columns = {"nama", "nip", "jabatan", "pangkat"}
        missing_columns = required_columns - set(df.columns)
        if missing_columns:
            raise ValueError(f"Kolom Excel tidak lengkap: {', '.join(sorted(missing_columns))}")

        df["nama"] = df["nama"].apply(normalisasi_teks)
        df["nip"] = df["nip"].apply(normalisasi_teks)
        df["jabatan"] = df["jabatan"].apply(lambda value: JABATAN_MAP.get(normalisasi_teks(value), normalisasi_teks(value)))
        df["pangkat_gol"] = df["pangkat"].apply(lambda value: PANGKAT_MAP.get(normalisasi_teks(value), normalisasi_teks(value)))
        df = df[df["nama"] != ""]

        return df.set_index("nama").to_dict(orient="index")
    except FileNotFoundError:
        st.error("File 'data_pegawai.xlsx' tidak ditemukan di folder aplikasi! Silakan periksa kembali.")
        return {}
    except Exception as error:
        st.error(f"Gagal membaca data Excel: {error}")
        return {}

DATA_PEGAWAI = muat_data_excel()

st.set_page_config(page_title="Sistem Merge Word & TTD Digital", layout="centered")

st.title("Formulir Administrasi Pegawai (Template Word)")
st.write("Silakan isi data di bawah ini untuk mengunduh dokumen Word yang telah terisi.")

# Jika data Excel berhasil dimuat
if DATA_PEGAWAI:
    # Buat daftar nama untuk dropdown, tambahkan opsi default di paling atas
    daftar_nama = ["Pilih Nama Peserta..."] + list(DATA_PEGAWAI.keys())

    # =========================================================================
    # MEKANISME AUTO-FILL
    # =========================================================================
    # 1. Pilih nama terlebih dahulu
    nama_pilihan = st.selectbox("Pilih Nama Anda", daftar_nama)

    # 2. Inisialisasi nilai default kosong jika belum memilih nama
    nip_val = ""
    index_pangkat = 8  # Default ke Penata Muda (III/a)
    index_jabatan = 2  # Default ke Guru Ahli Pertama

    # 3. Jika user sudah memilih nama asli (bukan opsi default)
    if nama_pilihan != "Pilih Nama Peserta...":
        data_terpilih = DATA_PEGAWAI[nama_pilihan]
        nip_val = data_terpilih.get("nip", "")
        
        # Cari posisi indeks pangkat berdasarkan data Excel
        pangkat_excel = data_terpilih.get("pangkat_gol", "")
        if pangkat_excel in LIST_PANGKAT:
            index_pangkat = LIST_PANGKAT.index(pangkat_excel)
            
        # Cari posisi indeks jabatan berdasarkan data Excel
        jabatan_excel = data_terpilih.get("jabatan", "")
        if jabatan_excel in LIST_JABATAN:
            index_jabatan = LIST_JABATAN.index(jabatan_excel)

    # =========================================================================
    # FORM UTAMA ADMINISTRASI
    # =========================================================================
    with st.form("form_administrasi_docx"):
        col1, col2 = st.columns(2)
        with col1:
            # Menampilkan nama terpilih (disabled agar tidak diubah manual)
            nama = st.text_input("Nama Lengkap", value="" if nama_pilihan == "Pilih Nama Peserta..." else nama_pilihan, disabled=True)
            
            # NIP otomatis terisi dari data Excel
            nip = st.text_input("NIP", value=nip_val, max_chars=18, placeholder="19xxxxxxxxxxxxxx")
            
        with col2:
            # Dropdown Pangkat otomatis bergeser sesuai data Excel
            pangkat_gol = st.selectbox("Pangkat / Golongan Ruang", LIST_PANGKAT, index=index_pangkat)
            
            # Dropdown Jabatan otomatis bergeser sesuai data Excel
            jabatan = st.selectbox("Jabatan", LIST_JABATAN, index=index_jabatan)

        # 2. Canvas Signature Pad untuk Tanda Tangan Digital
        st.markdown("---")
        st.subheader("Bubuhi Tanda Tangan Anda di Bawah Ini:")
        
        canvas_result = st_canvas(
            fill_color="rgba(255, 255, 255, 0)",
            stroke_width=3,
            stroke_color="#000000",
            background_color="#ffffff",
            height=150,
            width=400,
            drawing_mode="freedraw",
            key="canvas_ttd_docx",
        )
        
        submit_button = st.form_submit_button("Gabungkan ke Template Word")

    # 3. Proses Mail Merge menggunakan docxtpl
    if submit_button:
        if nama_pilihan == "Pilih Nama Peserta..." or not nip or not jabatan:
            st.error("Gagal! Silakan pilih Nama Anda dan pastikan NIP serta Jabatan sudah terisi.")
        elif canvas_result.image_data is None:
            st.error("Gagal! Tanda tangan tidak boleh kosong.")
        else:
            with st.spinner("Sedang menggabungkan data ke template Word..."):
                try:
                    doc = DocxTemplate("template_anda.docx")
                    
                    img_data = canvas_result.image_data
                    img = Image.fromarray(img_data.astype('uint8'), 'RGBA')
                    
                    img_buffer = io.BytesIO()
                    img.save(img_buffer, format="PNG")
                    img_buffer.seek(0)
                    
                    tanda_tangan_image = InlineImage(doc, img_buffer, width=Inches(1.5))
                    
                    context = {
                        'nama': nama_pilihan,
                        'nip': nip,
                        'pangkat_gol': pangkat_gol,
                        'jabatan': jabatan,  
                        'ttd': tanda_tangan_image,
                        'tanggal_sekarang': datetime.date.today().strftime('%d %B %Y') 
                    }
                    
                    doc.render(context)
                    
                    output_buffer = io.BytesIO()
                    doc.save(output_buffer)
                    output_buffer.seek(0)
                    
                    st.success(f"Dokumen untuk {nama_pilihan} berhasil dibuat!")
                    nama_file = re.sub(r'[\\/:*?"<>|]', "_", nama_pilihan)
                    st.download_button(
                        label="Unduh Dokumen Word (.docx)",
                        data=output_buffer.getvalue(),
                        file_name=f"Dokumen_Administrasi_{nama_file}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except FileNotFoundError:
                    st.error("File 'template_anda.docx' tidak ditemukan. Pastikan file template berada di direktori yang sama.")
else:
    st.warning("Aplikasi tidak dapat dijalankan karena data referensi Excel belum siap.")