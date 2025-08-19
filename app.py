import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from docx import Document
from io import BytesIO
import re
import json

# --- Konfigurasi Google Sheets ---
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
def process_docx(uploaded_file, data_map):
    try:
        doc_stream = BytesIO(uploaded_file.read())
        doc = Document(doc_stream)
        
        # Regex untuk menemukan tag dengan prefix dan cell, contoh: [Data:A1]
        pattern = re.compile(r"\[(.*?:[A-Z]+[0-9]+)\]")
        
        # Fungsi pembantu untuk memproses teks di dalam sebuah kontainer (paragraph atau cell)
        def replace_in_container(container):
            full_text = ""
            for run in container.runs:
                full_text += run.text
            
            matches = pattern.findall(full_text)
            
            if not matches:
                return

            # Hanya lakukan penggantian jika ada tag yang ditemukan
            for match in matches:
                replacement_value = str(data_map.get(match, f"[{match}]"))
                
                # Cari dan ganti tag di dalam full_text
                full_text = full_text.replace(f"[{match}]", replacement_value)

            # Hapus semua teks dari run yang ada
            for run in container.runs:
                run.text = ""
            
            # Masukkan teks yang sudah diganti ke dalam run pertama
            if container.runs:
                container.runs[0].text = full_text
            else:
                # Jika tidak ada run, tambahkan run baru
                container.add_run(full_text)
                
        # Memproses paragraf
        for p in doc.paragraphs:
            replace_in_container(p)
        
        # Memproses tabel
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    # Setiap cell memiliki paragraphs
                    for p in cell.paragraphs:
                        replace_in_container(p)
        
        new_doc_stream = BytesIO()
        doc.save(new_doc_stream)
        new_doc_stream.seek(0)
        
        return new_doc_stream
        
    except Exception as e:
        st.error(f"Error saat memproses dokumen: {e}")
        return None

# --- Antarmuka Streamlit ---
st.title("Aplikasi Otomatisasi Dokumen Word Fleksibel 🚀")
st.markdown("Aplikasi ini akan mengganti tag dalam dokumen Word (.docx) dengan data dari Google Sheet.")
st.markdown("---")

st.header("Konfigurasi Data Google Sheet")
st.markdown("Masukkan URL Google Sheet dan konfigurasi sheet dalam format JSON. Setiap entri adalah pasangan dari **'prefix tag'** dan **'nama sheet'**.")

sheet_url_input = st.text_input(
    "URL Google Sheet", 
    "https://docs.google.com/spreadsheets/d/1d89txS35ZrBwk6_gyfOixvWZlvc149pgzhbekOka1uo/edit?usp=sharing"
)

sheet_config_input = st.text_area(
    "Konfigurasi Sheet (JSON)", 
    '{\n    "Data": "Data",\n    "Obx": "Obx",\n    "Control": "Control"\n}', 
    height=150
)

uploaded_file = st.file_uploader("Pilih file .docx", type="docx")

if st.button("Generate Dokumen"):
    if not uploaded_file:
        st.warning("Silakan unggah file .docx terlebih dahulu.")
    else:
        with st.spinner("Sedang memproses..."):
            try:
                # 1. Parsing konfigurasi sheet dari input pengguna
                sheet_config = json.loads(sheet_config_input)
                
                # 2. Buat dictionary data gabungan dari semua sheet yang dikonfigurasi
                all_data_dict = {}
                for prefix, sheet_name in sheet_config.items():
                    try:
                        worksheet = client.open_by_url(sheet_url_input).worksheet(sheet_name)
                        all_values = worksheet.get_all_values()
                        
                        for r_idx, row in enumerate(all_values):
                            for c_idx, cell_value in enumerate(row):
                                col_letter = chr(ord('A') + c_idx)
                                cell_ref = f"{col_letter}{r_idx + 1}"
                                # Key dictionary adalah kombinasi dari prefix dan cell, contoh: "Data:A1"
                                combined_key = f"{prefix}:{cell_ref}"
                                all_data_dict[combined_key] = cell_value
                    except gspread.WorksheetNotFound:
                        st.warning(f"Sheet dengan nama '{sheet_name}' tidak ditemukan. Melewati sheet ini.")
                        
                if not all_data_dict:
                    st.error("Tidak ada data yang berhasil diambil dari Google Sheets. Mohon periksa URL dan nama sheet Anda.")
                else:
                    # 3. Proses dokumen Word dengan dictionary data gabungan
                    processed_doc_stream = process_docx(uploaded_file, all_data_dict)
                    
                    if processed_doc_stream:
                        file_name = uploaded_file.name.replace(".docx", "_generated.docx")
                        
                        st.success("Dokumen berhasil dibuat!")
                        st.download_button(
                            label="Unduh Dokumen Hasil",
                            data=processed_doc_stream,
                            file_name=file_name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )

            except json.JSONDecodeError:
                st.error("Konfigurasi Sheet tidak valid. Mohon gunakan format JSON yang benar.")
            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")
