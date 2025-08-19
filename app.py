import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from docx import Document
from io import BytesIO
import re

# --- Konfigurasi Google Sheets ---
# Pastikan file .streamlit/secrets.toml sudah disiapkan dengan benar.
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

try:
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scope
    )
    client = gspread.authorize(creds)
except Exception as e:
    st.error(f"Error saat otentikasi ke Google Sheets: {e}")
    st.stop()

# --- Fungsi untuk memproses dokumen Word ---
def process_docx(uploaded_file, data_map, prefix_tag):
    try:
        doc_stream = BytesIO(uploaded_file.read())
        doc = Document(doc_stream)
        
        # Regex untuk menemukan tag dengan prefix dinamis, contoh: [Data:A1], [Laporan:B2]
        pattern = re.compile(r"\[" + re.escape(prefix_tag) + r":([A-Z]+[0-9]+)\]")
        
        # Memproses paragraf
        for p in doc.paragraphs:
            for run in p.runs:
                matches = pattern.findall(run.text)
                if matches:
                    for match in matches:
                        replacement_value = str(data_map.get(match, f"[{prefix_tag}:{match}]"))
                        run.text = run.text.replace(f"[{prefix_tag}:{match}]", replacement_value)
        
        # Memproses tabel
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            matches = pattern.findall(run.text)
                            if matches:
                                for match in matches:
                                    replacement_value = str(data_map.get(match, f"[{prefix_tag}:{match}]"))
                                    run.text = run.text.replace(f"[{prefix_tag}:{match}]", replacement_value)
        
        new_doc_stream = BytesIO()
        doc.save(new_doc_stream)
        new_doc_stream.seek(0)
        
        return new_doc_stream
        
    except Exception as e:
        st.error(f"Error saat memproses dokumen: {e}")
        return None

# --- Antarmuka Streamlit ---
st.title("Aplikasi Otomatisasi Dokumen Word Fleksibel 📄")
st.markdown("Aplikasi ini akan mengganti tag dalam dokumen Word (.docx) dengan data dari Google Sheet.")
st.markdown("---")

# Input pengguna untuk URL, nama sheet, dan tag
st.header("Konfigurasi Data Google Sheet")
sheet_url_input = st.text_input(
    "URL Google Sheet", 
    "https://docs.google.com/spreadsheets/d/1d89txS35ZrBwk6_gyfOixvWZlvc149pgzhbekOka1uo/edit?usp=sharing"
)
sheet_name_input = st.text_input("Nama Sheet", "Data")
prefix_tag_input = st.text_input("Nama Prefix Tag", "Data")

st.markdown("---")

st.header("Unggah dan Proses Dokumen")
uploaded_file = st.file_uploader("Pilih file .docx", type="docx")

if st.button("Generate Dokumen"):
    if not uploaded_file:
        st.warning("Silakan unggah file .docx terlebih dahulu.")
    else:
        with st.spinner("Sedang memproses..."):
            try:
                # Buka sheet berdasarkan URL dan nama sheet dari input pengguna
                sheet = client.open_by_url(sheet_url_input)
                worksheet = sheet.worksheet(sheet_name_input)
                
                # Ambil semua nilai dari sheet
                all_values = worksheet.get_all_values()
                
                # Buat dictionary untuk mapping cell ke value
                data_dict = {}
                for r_idx, row in enumerate(all_values):
                    for c_idx, cell_value in enumerate(row):
                        col_letter = chr(ord('A') + c_idx)
                        cell_ref = f"{col_letter}{r_idx + 1}"
                        data_dict[cell_ref] = cell_value

                # Panggil fungsi untuk memproses dokumen
                processed_doc_stream = process_docx(uploaded_file, data_dict, prefix_tag_input)
                
                if processed_doc_stream:
                    file_name = uploaded_file.name.replace(".docx", "_generated.docx")
                    
                    st.success("Dokumen berhasil dibuat!")
                    st.download_button(
                        label="Unduh Dokumen Hasil",
                        data=processed_doc_stream,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            except gspread.WorksheetNotFound:
                st.error(f"Sheet dengan nama '{sheet_name_input}' tidak ditemukan. Mohon periksa kembali nama sheet Anda.")
            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")
