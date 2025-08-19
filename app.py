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
        
        pattern = re.compile(r"\[(.*?:[A-Z]+[0-9]+)\]")
        
        def replace_in_container(container):
            full_text = ""
            for run in container.runs:
                full_text += run.text
            
            matches = pattern.findall(full_text)
            
            if not matches:
                return

            for match in matches:
                replacement_value = str(data_map.get(match, f"[{match}]"))
                full_text = full_text.replace(f"[{match}]", replacement_value)

            for run in container.runs:
                run.text = ""
            
            if container.runs:
                container.runs[0].text = full_text
            else:
                container.add_run(full_text)
                
        for p in doc.paragraphs:
            replace_in_container(p)
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
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

# --- Tabs ---
tab1, tab2 = st.tabs(["Aplikasi", "Panduan Penggunaan"])

with tab1:
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

    st.markdown("---")

    st.header("Unggah dan Buat Dokumen")
    uploaded_file = st.file_uploader("Pilih file .docx", type="docx")

    if st.button("Generate Dokumen"):
        if not uploaded_file:
            st.warning("Silakan unggah file .docx terlebih dahulu.")
        else:
            with st.spinner("Sedang memproses..."):
                try:
                    sheet_config = json.loads(sheet_config_input)
                    all_data_dict = {}
                    for prefix, sheet_name in sheet_config.items():
                        try:
                            worksheet = client.open_by_url(sheet_url_input).worksheet(sheet_name)
                            all_values = worksheet.get_all_values()
                            for r_idx, row in enumerate(all_values):
                                for c_idx, cell_value in enumerate(row):
                                    col_letter = chr(ord('A') + c_idx)
                                    cell_ref = f"{col_letter}{r_idx + 1}"
                                    combined_key = f"{prefix}:{cell_ref}"
                                    all_data_dict[combined_key] = cell_value
                        except gspread.WorksheetNotFound:
                            st.warning(f"Sheet dengan nama '{sheet_name}' tidak ditemukan. Melewati sheet ini.")
                            
                    if not all_data_dict:
                        st.error("Tidak ada data yang berhasil diambil dari Google Sheets. Mohon periksa URL dan nama sheet Anda.")
                    else:
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

with tab2:
    st.header("🚀 Panduan Singkat: Mengisi Dokumen Word Otomatis")
    st.markdown("Aplikasi ini membantu Anda mengisi data dari Google Sheets ke dalam dokumen Word secara otomatis. Cukup ikuti 3 langkah mudah ini.")

    st.subheader("Langkah 1: Siapkan Google Sheets & Dokumen Word")
    st.markdown("1. **Buka Google Sheet Anda**")
    st.markdown(f"   - Kunjungi https://docs.google.com/spreadsheets/d/1d89txS35ZrBwk6_gyfOixvWZlvc149pgzhbekOka1uo/edit?usp=sharing untuk melihat contoh data.")
    st.markdown("   - Pastikan **akun layanan** Anda memiliki akses 'Editor' ke Google Sheet ini.")
    st.markdown("2. **Buat Template Word Anda**")
    st.markdown("   - Tulis tag di dokumen Word Anda dengan format **`[prefix:cell]`**.")
    st.markdown("   - **`prefix`**: Nama yang Anda tentukan di aplikasi untuk mewakili sebuah sheet.")
    st.markdown("   - **`cell`**: Nama sel yang ingin Anda ambil datanya (contoh: `A1`, `B2`).")
    st.markdown("**Contoh Tag**: `[Data:A1]`, `[Obx:A2]`, `[Control:A3]`")

    st.markdown("---")

    st.subheader("Langkah 2: Konfigurasi di Aplikasi Web")
    st.markdown("1. **Buka Aplikasi**")
    st.markdown("   - Kunjungi **https://exceltowordz.streamlit.app/**.")
    st.markdown("   - Pilih tab **'Aplikasi'**.")
    st.markdown("2. **Atur Konfigurasi JSON**")
    st.markdown("   - Masukkan URL Google Sheet Anda.")
    st.markdown("   - Pada bagian **'Konfigurasi Sheet'**, masukkan nama sheet yang ingin Anda gunakan dalam format JSON. Setiap pasangan `“prefix”: “nama_sheet”` memberitahu aplikasi sheet mana yang harus diakses untuk prefix tertentu.")
    st.code("{\n  \"Data\": \"Data\",\n  \"Obx\": \"Obx\",\n  \"Control\": \"Control\"\n}")
    st.markdown("Jika Anda ingin **menambah sheet baru**, cukup tambahkan baris baru ke dalam JSON. Misalnya, untuk sheet 'Finance' dengan tag `[Finance:B1]`, tambahkan:")
    st.code("{\n  \"Data\": \"Data\",\n  \"Obx\": \"Obx\",\n  \"Control\": \"Control\",\n  \"Finance\": \"Finance\"\n}")
    st.markdown("Catatan: Nama sel (`A1`, `B2`, dll.) **tidak perlu** dimasukkan di sini.")

    st.markdown("---")

    st.subheader("Langkah 3: Unggah & Unduh Dokumen")
    st.markdown("1. **Unggah Dokumen**")
    st.markdown("   - Klik tombol **'Browse files'** dan pilih file Word (.docx) template Anda.")
    st.markdown("2. **Generate**")
    st.markdown("   - Klik tombol **'Generate Dokumen'**. Aplikasi akan memproses file Anda. Tunggu sebentar hingga proses selesai.")
    st.markdown("3. **Unduh**")
    st.markdown("   - Setelah proses berhasil, tombol **'Unduh Dokumen Hasil'** akan muncul.")
    st.markdown("   - Klik tombol tersebut untuk mengunduh dokumen Word yang sudah terisi penuh dengan data dari Google Sheets Anda.")
