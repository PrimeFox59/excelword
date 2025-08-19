import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from docx import Document
from io import BytesIO
import re

# --- Konfigurasi Google Sheets ---
# Ganti dengan link Google Sheet Anda dan secret key dari file secret.toml
# Untuk deployment, secret key ini harus disimpan di Streamlit Secrets
# dengan nama file .streamlit/secrets.toml
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

try:
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scope
    )
    client = gspread.authorize(creds)

    sheet_url = "https://docs.google.com/spreadsheets/d/1d89txS35ZrBwk6_gyfOixvWZlvc149pgzhbekOka1uo/edit?usp=sharing"
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.worksheet("Data")

    # Ambil semua nilai dari sheet
    all_values = worksheet.get_all_values()

    # Buat dictionary untuk mapping cell ke value
    data_dict = {}
    for r_idx, row in enumerate(all_values):
        for c_idx, cell_value in enumerate(row):
            col_letter = chr(ord('A') + c_idx)
            cell_ref = f"{col_letter}{r_idx + 1}"
            data_dict[cell_ref] = cell_value
except Exception as e:
    st.error(f"Error saat terhubung ke Google Sheets: {e}")
    st.stop()

# --- Fungsi untuk memproses dokumen Word ---
def process_docx(uploaded_file, data_map):
    try:
        # Buka dokumen yang diunggah dari buffer memori
        doc_stream = BytesIO(uploaded_file.read())
        doc = Document(doc_stream)
        
        # Regex untuk menemukan tag seperti [Data:A1], [Data:B2], dll.
        pattern = re.compile(r"\[Data:([A-Z]+[0-9]+)\]")
        
        for p in doc.paragraphs:
            for run in p.runs:
                matches = pattern.findall(run.text)
                if matches:
                    for match in matches:
                        # Dapatkan nilai dari data_map
                        # Jika cell tidak ditemukan, biarkan tag asli atau ganti dengan string kosong
                        replacement_value = str(data_map.get(match, f"[Data:{match}]"))
                        run.text = run.text.replace(f"[Data:{match}]", replacement_value)
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            matches = pattern.findall(run.text)
                            if matches:
                                for match in matches:
                                    replacement_value = str(data_map.get(match, f"[Data:{match}]"))
                                    run.text = run.text.replace(f"[Data:{match}]", replacement_value)
        
        # Simpan dokumen yang sudah dimodifikasi ke buffer memori
        new_doc_stream = BytesIO()
        doc.save(new_doc_stream)
        new_doc_stream.seek(0)
        
        return new_doc_stream
        
    except Exception as e:
        st.error(f"Error saat memproses dokumen: {e}")
        return None

# --- Antarmuka Streamlit ---
st.title("Aplikasi Otomatisasi Dokumen Word")
st.markdown("Unggah file Word Anda (.docx) yang berisi tag seperti `[Data:A1]` atau `[Data:C5]` untuk diganti dengan data dari Google Sheet.")
st.markdown("---")

uploaded_file = st.file_uploader("Pilih file .docx", type="docx")

if uploaded_file:
    st.info("File berhasil diunggah. Klik 'Generate' untuk memproses dokumen.")
    
    # Tombol Generate
    if st.button("Generate"):
        st.spinner("Sedang memproses dokumen...")
        
        # Panggil fungsi untuk memproses dokumen
        processed_doc_stream = process_docx(uploaded_file, data_dict)
        
        if processed_doc_stream:
            # Dapatkan nama file asli untuk diunduh
            file_name = uploaded_file.name.replace(".docx", "_generated.docx")
            
            # Tampilkan tombol unduh
            st.success("Dokumen berhasil dibuat!")
            st.download_button(
                label="Unduh Dokumen Hasil",
                data=processed_doc_stream,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
else:
    st.warning("Silakan unggah file .docx terlebih dahulu.")