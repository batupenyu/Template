import streamlit as st
from streamlit_drawable_canvas import st_canvas
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches  # Mengatur ukuran gambar di Word
from PIL import Image
import io
import datetime

st.set_page_config(page_title="Sistem Merge Word & TTD Digital", layout="centered")

st.title("Formulir Administrasi Pegawai (Template Word)")
st.write("Silakan isi data di bawah ini untuk mengunduh dokumen Word yang telah terisi.")

# 1. Input Field Administrasi (Termasuk Field Jabatan Baru)
with st.form("form_administrasi_docx"):
    col1, col2 = st.columns(2)
    with col1:
        nama = st.text_input("Nama Lengkap", placeholder="Nama beserta gelar")
        nip = st.text_input("NIP", max_chars=18, placeholder="19xxxxxxxxxxxxxx")
    with col2:
        pangkat_gol = st.selectbox("Pangkat / Golongan Ruang", [
            "Pembina Utama (IV/e)", "Pembina Utama Madya (IV/d)", "Pembina Utama Muda (IV/c)",
            "Pembina Tingkat I (IV/b)", "Pembina (IV/a)", "Penata Tingkat I (III/d)",
            "Penata (III/c)", "Penata Muda Tingkat I (III/b)", "Penata Muda (III/a)"
        ])
        jabatan = st.selectbox("Jabatan", [
            "Guru Ahli Muda",
            "Guru Ahli Pertama",
            "Guru Ahli Madya",
            "Bendahara",
            "Analis Sumber Daya Manusia Aparatur Ahli Muda",
        ])

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
    if not nama or not nip or not jabatan:
        st.error("Gagal! Nama, NIP, dan Jabatan wajib diisi.")
    elif canvas_result.image_data is None:
        st.error("Gagal! Tanda tangan tidak boleh kosong.")
    else:
        with st.spinner("Sedang menggabungkan data ke template Word..."):
            try:
                # Muat file template .docx Anda (pastikan file berada di folder yang sama atau tulis path lengkapnya)
                doc = DocxTemplate("template_anda.docx")
                
                # Proses gambar tanda tangan dari canvas
                img_data = canvas_result.image_data
                img = Image.fromarray(img_data.astype('uint8'), 'RGBA')
                
                # Simpan gambar sementara ke memori buffer (tanpa mengotori penyimpanan lokal)
                img_buffer = io.BytesIO()
                img.save(img_buffer, format="PNG")
                img_buffer.seek(0)
                
                # Atur tanda tangan agar bisa disisipkan ke dalam baris Word (Lebar disetel 1.5 inci)
                tanda_tangan_image = InlineImage(doc, img_buffer, width=Inches(1.5))
                
                # Satukan semua field ke dalam dictionary context
                context = {
                    'nama': nama,
                    'nip': nip,
                    'pangkat_gol': pangkat_gol,
                    'jabatan': jabatan,  # Field baru berhasil dipetakan
                    'ttd': tanda_tangan_image,
                    'tanggal_sekarang': datetime.date.today().strftime('%d %B %Y') # Opsional jika butuh tag tanggal otomatis
                }
                
                # Jalankan perintah render / penggabungan
                doc.render(context)
                
                # Simpan output dokumen hasil ke dalam memori buffer
                output_buffer = io.BytesIO()
                doc.save(output_buffer)
                output_buffer.seek(0)
                
                st.success("Dokumen .docx berhasil digenerate!")
                st.download_button(
                    label="Unduh Dokumen Word (.docx)",
                    data=output_buffer.getvalue(),
                    file_name=f"Dokumen_Administrasi_{nip}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except FileNotFoundError:
                st.error("File 'template_anda.docx' tidak ditemukan. Pastikan file template berada di direktori yang sama dengan skrip ini.")