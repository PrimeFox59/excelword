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
        
        # Regex untuk menemukan tag dengan prefix dinamis, contoh: [Data:A1], [Obx:B2]
        # Regex ini akan mengambil seluruh tag, termasuk prefix dan cell
        pattern = re.compile(r"\[(.*?:[A-Z]+[0-9]+)\]")
        
        # Memproses paragraf
        for p in doc.paragraphs:
            runs_to_replace = []
            for run in p.runs:
                matches = pattern.findall(run.text)
                if matches:
                    for match in matches:
                        runs_to_replace.append((run, match))

            # Lakukan penggantian setelah mengumpulkan semua match
            for run, match in runs_to_replace:
                replacement_value = str(data_map.get(match, f"[{match}]"))
                run.text = run.text.replace(f"[{match}]", replacement_value)
        
        # Memproses tabel
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    runs_to_replace_in_cell = []
                    for p in cell.paragraphs:
                        for run in p.runs:
                            matches = pattern.findall(run.text)
                            if matches:
                                for match in matches:
                                    runs_to_replace_in_cell.append((run, match))

                    for run, match in runs_to_replace_in_cell:
                        replacement_value = str(data_map.get(match, f"[{match}]"))
                        run.text = run.text.replace(f"[{match}]", replacement_value)
        
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
