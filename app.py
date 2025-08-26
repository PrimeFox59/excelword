import streamlit as st
import pandas as pd
from datetime import datetime
from fpdf import FPDF
import io
import bcrypt
import plotly.express as px

# --- Google Sheets Integration ---
from google.oauth2.service_account import Credentials
import gspread

# Placeholder for user to add their sheet name
GOOGLE_SHEET_NAME = "PT. BERKAT KARYA ANUGERAH"

def get_client():
    try:
        # Gunakan st.secrets untuk otentikasi
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], scopes=scope
        )
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"Error saat otentikasi Google Sheets: {e}.")
        st.info("Pastikan Anda telah mengonfigurasi `gcp_service_account` di Streamlit secrets. Mode Google Sheets tidak akan berfungsi.")
        return None

# --- CUSTOM CSS FOR UI/UX IMPROVEMENTS ---
st.markdown("""
<style>
    .reportview-container {
        background: #F0F2F6;
    }
    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }
    .st-emotion-cache-1r6509j {
        background-color: #2F3E50;
    }
    .st-emotion-cache-1r6509j .stButton>button {
        color: white;
    }
    .st-emotion-cache-1r6509j .stButton>button:hover {
        background-color: #455A64;
    }
    .st-emotion-cache-1r6509j .st-bv {
        color: white;
    }
    .st-emotion-cache-16k1w7m {
        background-color: #2F3E50;
    }
    h1, h2, h3, h4, h5, h6 {
        color: #2F3E50;
    }
    .stTabs [data-baseweb="tab-list"] button [data-testid="stMarkdownContainer"] p {
        font-size: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# --- Database & Sheets Interaction Functions ---
@st.cache_data(show_spinner=False)
def get_data_from_sheet(sheet_name):
    client = get_client()
    if client:
        try:
            sheet = client.open(GOOGLE_SHEET_NAME).worksheet(sheet_name)
            data = sheet.get_all_records()
            df = pd.DataFrame(data)
            return df
        except Exception as e:
            st.error(f"Error saat membaca data dari sheet '{sheet_name}'. Pastikan sheet dengan nama tersebut ada. Error: {e}")
            return pd.DataFrame()
    return pd.DataFrame()

def write_to_sheet(sheet_name, data_list):
    client = get_client()
    if client:
        try:
            sheet = client.open(GOOGLE_SHEET_NAME).worksheet(sheet_name)
            sheet.append_row(data_list)
            st.cache_data.clear()
            return True
        except Exception as e:
            st.error(f"Error saat menulis ke sheet '{sheet_name}'. Error: {e}")
            return False
    return False

def update_sheet_row(sheet_name, row_index, data_list):
    client = get_client()
    if client:
        try:
            sheet = client.open(GOOGLE_SHEET_NAME).worksheet(sheet_name)
            sheet.update(f'A{row_index+2}', [data_list]) # gspread update row is 1-indexed, headers are A1, so A2 is the first data row
            st.cache_data.clear()
            return True
        except Exception as e:
            st.error(f"Error saat memperbarui row di sheet '{sheet_name}'. Error: {e}")
            return False
    return False

def delete_sheet_row(sheet_name, row_index):
    client = get_client()
    if client:
        try:
            sheet = client.open(GOOGLE_SHEET_NAME).worksheet(sheet_name)
            sheet.delete_rows(row_index + 2) # gspread delete is 1-indexed, headers are A1, so A2 is the first data row
            st.cache_data.clear()
            return True
        except Exception as e:
            st.error(f"Error saat menghapus row di sheet '{sheet_name}'. Error: {e}")
            return False
    return False
    
# --- Authentication & Initial Setup ---
def check_login(username, password):
    users_df = get_data_from_sheet('users')
    user_row = users_df[users_df['username'] == username]
    if not user_row.empty:
        hashed_password = user_row.iloc[0]['password_hash'].encode('utf-8')
        return bcrypt.checkpw(password.encode('utf-8'), hashed_password)
    return False

def init_app_state():
    # Only need to check for the 'owner' user, the sheets are assumed to exist
    users_df = get_data_from_sheet('users')
    
    # Check if the 'users' sheet is empty or if the 'owner' user does not exist
    if users_df.empty or 'owner' not in users_df['username'].tolist():
        st.warning("Data sheet 'users' tidak ditemukan atau user 'owner' tidak ada. Mencoba membuat user 'owner'...")
        hashed_password = bcrypt.hashpw("owner123".encode('utf-8'), bcrypt.gensalt())
        if write_to_sheet('users', ['owner', hashed_password.decode('utf-8')]):
            st.success("User 'owner' dengan password 'owner123' berhasil ditambahkan. Silakan login.")
            st.info("Anda mungkin perlu me-refresh halaman jika tidak ada tombol login.")
        else:
            st.error("Gagal menambahkan user 'owner'. Mohon periksa kembali konfigurasi Google Sheets.")

    # Check if the existing owner's password hash is valid
    else:
        owner_row = users_df[users_df['username'] == 'owner']
        if not owner_row.empty:
            hashed_password = owner_row.iloc[0]['password_hash']
            try:
                # Validate the hash by trying to decode it
                if not isinstance(hashed_password, str) or not hashed_password.startswith('$2'):
                    st.warning("Password hash untuk user 'owner' tidak valid. Mencoba membuat ulang...")
                    hashed_password = bcrypt.hashpw("owner123".encode('utf-8'), bcrypt.gensalt())
                    
                    # Update the existing row instead of adding a new one
                    row_index = owner_row.index[0]
                    client = get_client()
                    if client:
                        sheet = client.open(GOOGLE_SHEET_NAME).worksheet('users')
                        sheet.update(f'B{row_index+2}', hashed_password.decode('utf-8'))
                        st.success("Password hash user 'owner' berhasil diperbarui.")
                    else:
                        st.error("Gagal memperbarui password hash. Tidak dapat terhubung ke Google Sheets.")
            except Exception as e:
                st.error(f"Terjadi kesalahan saat validasi hash: {e}")

# --- CRUD Functions - Inventory (using sheets) ---
def add_master_item(kode, supplier, nama, warna, rak, harga):
    master_df = get_data_from_sheet('master_barang')
    if not master_df.empty and ((master_df['kode_bahan'] == kode) & (master_df['warna'] == warna)).any():
        return False
    return write_to_sheet('master_barang', [kode, supplier, nama, warna, rak, harga])

def get_master_barang():
    return get_data_from_sheet('master_barang')

def update_master_item(old_kode, old_warna, new_kode, new_warna, supplier, nama, rak, harga):
    master_df = get_master_barang()
    if (master_df['kode_bahan'] == old_kode).any() and (master_df['warna'] == old_warna).any():
        row_index = master_df[(master_df['kode_bahan'] == old_kode) & (master_df['warna'] == old_warna)].index[0]
        if (new_kode != old_kode or new_warna != old_warna):
             if ((master_df['kode_bahan'] == new_kode) & (master_df['warna'] == new_warna)).any():
                 return False
        return update_sheet_row('master_barang', row_index, [new_kode, supplier, nama, new_warna, rak, harga])
    return False

def delete_master_item(kode, warna):
    master_df = get_master_barang()
    if (master_df['kode_bahan'] == kode).any() and (master_df['warna'] == warna).any():
        row_index = master_df[(master_df['kode_bahan'] == kode) & (master_df['warna'] == warna)].index[0]
        return delete_sheet_row('master_barang', row_index)
    return False

def add_barang_masuk(tanggal_waktu, kode_bahan, warna, stok, yard, keterangan):
    barang_masuk_df = get_data_from_sheet('barang_masuk')
    new_id = barang_masuk_df['id'].max() + 1 if not barang_masuk_df.empty else 1
    return write_to_sheet('barang_masuk', [new_id, tanggal_waktu, kode_bahan, warna, stok, yard, keterangan])

def get_barang_masuk():
    return get_data_from_sheet('barang_masuk')

def update_barang_masuk(id, tanggal_waktu, kode_bahan, warna, stok, yard, keterangan):
    barang_masuk_df = get_barang_masuk()
    if (barang_masuk_df['id'] == id).any():
        row_index = barang_masuk_df[barang_masuk_df['id'] == id].index[0]
        return update_sheet_row('barang_masuk', row_index, [id, tanggal_waktu, kode_bahan, warna, stok, yard, keterangan])
    return False

def delete_barang_masuk(id):
    barang_masuk_df = get_barang_masuk()
    if (barang_masuk_df['id'] == id).any():
        row_index = barang_masuk_df[barang_masuk_df['id'] == id].index[0]
        return delete_sheet_row('barang_masuk', row_index)
    return False

@st.cache_data(ttl=600)
def get_stock_balance(kode_bahan, warna):
    barang_masuk_df = get_data_from_sheet('barang_masuk')
    barang_keluar_df = get_data_from_sheet('barang_keluar')
    
    in_stock = barang_masuk_df[(barang_masuk_df['kode_bahan'] == kode_bahan) & (barang_masuk_df['warna'] == warna)]['stok'].sum()
    out_stock = barang_keluar_df[(barang_keluar_df['kode_bahan'] == kode_bahan) & (barang_keluar_df['warna'] == warna)]['stok'].sum()

    return in_stock - out_stock

def get_in_out_records(start_date, end_date):
    in_df = get_data_from_sheet('barang_masuk')
    out_df = get_data_from_sheet('barang_keluar')

    if not in_df.empty:
        in_df['tanggal_waktu'] = pd.to_datetime(in_df['tanggal_waktu']).dt.date.astype(str)
        in_df = in_df[(in_df['tanggal_waktu'] >= start_date) & (in_df['tanggal_waktu'] <= end_date)]
        in_df['qty'] = in_df['stok']
        in_df['type'] = 'Masuk'
    else:
        in_df = pd.DataFrame(columns=['tanggal_waktu', 'kode_bahan', 'warna', 'qty', 'type', 'keterangan'])
    
    if not out_df.empty:
        out_df['tanggal_waktu'] = pd.to_datetime(out_df['tanggal_waktu']).dt.date.astype(str)
        out_df = out_df[(out_df['tanggal_waktu'] >= start_date) & (out_df['tanggal_waktu'] <= end_date)]
        out_df['qty'] = out_df['stok']
        out_df['type'] = 'Keluar'
    else:
        out_df = pd.DataFrame(columns=['tanggal_waktu', 'kode_bahan', 'warna', 'qty', 'type', 'keterangan'])

    df = pd.concat([in_df[['tanggal_waktu', 'kode_bahan', 'warna', 'qty', 'type', 'keterangan']], out_df[['tanggal_waktu', 'kode_bahan', 'warna', 'qty', 'type', 'keterangan']]], ignore_index=True)
    df['tanggal_waktu'] = pd.to_datetime(df['tanggal_waktu'])
    df = df.sort_values(by='tanggal_waktu')
    return df

# --- Invoice Functions ---
def get_invoices():
    df = get_data_from_sheet('invoices')
    if not df.empty:
        df['tanggal_waktu'] = pd.to_datetime(df['tanggal_waktu'])
        df = df.sort_values(by='tanggal_waktu', ascending=False)
    return df

def get_invoice_items(invoice_number):
    df = get_data_from_sheet('invoice_items')
    return df[df['invoice_number'] == invoice_number]

def get_barang_keluar():
    return get_data_from_sheet('barang_keluar')

def generate_invoice_pdf(invoice_data, invoice_items):
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    
    pdf.cell(0, 10, 'PT. BERKAT KARYA ANUGERAH', 0, 1, 'C')
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, 'INVOICE', 0, 1, 'C')
    pdf.set_font("Arial", '', 12)
    pdf.ln(5)
    
    pdf.cell(0, 5, f"No Invoice: {invoice_data['No Invoice']}", 0, 1, 'L')
    pdf.cell(0, 5, f"Tanggal: {invoice_data['Tanggal & Waktu']}", 0, 1, 'L')
    pdf.cell(0, 5, f"Nama Pelanggan: {invoice_data['Nama Pelanggan']}", 0, 1, 'L')
    
    pdf.ln(10)

    pdf.set_font("Arial", 'B', 12)
    pdf.cell(10, 10, 'No', 1, 0, 'C')
    pdf.cell(70, 10, 'Item', 1, 0, 'C')
    pdf.cell(30, 10, 'Qty', 1, 0, 'C')
    pdf.cell(40, 10, 'Harga', 1, 0, 'C')
    pdf.cell(40, 10, 'Total Harga', 1, 1, 'C')

    pdf.set_font("Arial", '', 12)
    total_invoice_amount = 0
    for idx, row in invoice_items.iterrows():
        total_invoice_amount += row['total']
        pdf.cell(10, 10, str(idx + 1), 1, 0, 'C')
        pdf.cell(70, 10, row['nama_bahan'], 1)
        pdf.cell(30, 10, str(row['qty']), 1, 0, 'R')
        pdf.cell(40, 10, f"Rp {row['harga']:,.2f}", 1, 0, 'R')
        pdf.cell(40, 10, f"Rp {row['total']:,.2f}", 1, 1, 'R')

    pdf.set_font("Arial", 'B', 12)
    pdf.cell(150, 10, 'Total', 1, 0, 'R')
    pdf.cell(40, 10, f"Rp {total_invoice_amount:,.2f}", 1, 1, 'R')
    
    pdf.ln(10)
    pdf.set_font("Arial", '', 12)
    pdf.cell(0, 5, "Terimakasih atas pembelian anda", 0, 1, 'C')
    pdf.ln(10)
    pdf.cell(0, 5, "Ttd Accounting", 0, 1, 'R')
    
    return pdf.output(dest='S').encode('latin1')
    
def generate_invoice_number():
    invoices_df = get_invoices()
    today_date = datetime.now().strftime('%y%m%d')
    prefix = f"INV-{today_date}-"
    
    last_invoice = invoices_df[invoices_df['invoice_number'].str.startswith(prefix)]['invoice_number'].max()
    
    if pd.notna(last_invoice):
        last_seq = int(last_invoice.split('-')[-1])
        new_seq = last_seq + 1
    else:
        new_seq = 1
        
    new_invoice_number = f"{prefix}{new_seq:03d}"
    return new_invoice_number

def add_barang_keluar_and_invoice(invoice_number, customer_name, items):
    tanggal_waktu = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    invoice_data_list = [invoice_number, tanggal_waktu, customer_name]
    if not write_to_sheet('invoices', invoice_data_list):
        return False, "Gagal mencatat invoice utama."

    for item in items:
        current_stock = get_stock_balance(item['kode_bahan'], item['warna'])
        if item['qty'] > current_stock:
            return False, f"Stok untuk bahan {item['nama_bahan']} ({item['warna']}) tidak mencukupi. Stok saat ini: {current_stock}"

        # Insert into invoice_items sheet
        invoice_items_df = get_data_from_sheet('invoice_items')
        new_item_id = invoice_items_df['id'].max() + 1 if not invoice_items_df.empty else 1
        invoice_item_list = [new_item_id, invoice_number, item['kode_bahan'], item['nama_bahan'], item['qty'], item['harga'], item['total']]
        if not write_to_sheet('invoice_items', invoice_item_list):
            return False, "Gagal mencatat item invoice."

        # Insert into barang_keluar sheet
        barang_keluar_df = get_data_from_sheet('barang_keluar')
        new_keluar_id = barang_keluar_df['id'].max() + 1 if not barang_keluar_df.empty else 1
        keluar_data_list = [new_keluar_id, tanggal_waktu, item['kode_bahan'], item['warna'], item['qty'], item['yard'], item['keterangan']]
        if not write_to_sheet('barang_keluar', keluar_data_list):
            return False, "Gagal mencatat barang keluar."
    
    return True, "Transaksi berhasil dicatat dan invoice dibuat."


# --- Payroll Functions ---
def add_employee(nama, bagian, gaji):
    employees_df = get_data_from_sheet('employees')
    new_id = employees_df['id'].max() + 1 if not employees_df.empty else 1
    return write_to_sheet('employees', [new_id, nama, bagian, gaji])

def get_employees():
    return get_data_from_sheet('employees')

def update_employee(employee_id, nama, bagian, gaji):
    employees_df = get_employees()
    if (employees_df['id'] == employee_id).any():
        row_index = employees_df[employees_df['id'] == employee_id].index[0]
        return update_sheet_row('employees', row_index, [employee_id, nama, bagian, gaji])
    return False

def delete_employee(employee_id):
    employees_df = get_employees()
    if (employees_df['id'] == employee_id).any():
        row_index = employees_df[employees_df['id'] == employee_id].index[0]
        return delete_sheet_row('employees', row_index)
    return False

def add_payroll_record(employee_id, gaji_bulan, gaji_pokok, lembur, lembur_minggu, uang_makan, pot_absen_finger, ijin_hr, simpanan_wajib, potongan_koperasi, kasbon, gaji_akhir, keterangan):
    payroll_df = get_data_from_sheet('payroll')
    new_id = payroll_df['id'].max() + 1 if not payroll_df.empty else 1
    tanggal_waktu = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    data_list = [new_id, tanggal_waktu, gaji_bulan, employee_id, gaji_pokok, lembur, lembur_minggu, uang_makan, pot_absen_finger, ijin_hr, simpanan_wajib, potongan_koperasi, kasbon, gaji_akhir, keterangan]
    return write_to_sheet('payroll', data_list)

def get_payroll_records():
    payroll_df = get_data_from_sheet('payroll')
    employees_df = get_data_from_sheet('employees')

    if payroll_df.empty or employees_df.empty:
        return pd.DataFrame()

    df = pd.merge(payroll_df, employees_df[['id', 'nama_karyawan']], left_on='employee_id', right_on='id', suffixes=('', '_emp'))
    df = df.sort_values(by='tanggal_waktu', ascending=False)
    return df[['id', 'tanggal_waktu', 'gaji_bulan', 'nama_karyawan', 'gaji_akhir', 'keterangan']]

def get_payroll_records_by_month(month_str):
    payroll_df = get_data_from_sheet('payroll')
    employees_df = get_data_from_sheet('employees')

    if payroll_df.empty or employees_df.empty:
        return pd.DataFrame()

    df = pd.merge(payroll_df, employees_df, left_on='employee_id', right_on='id', suffixes=('', '_emp'))
    df = df[df['gaji_bulan'] == month_str]
    return df

# --- Helper function to get unique yards ---
def get_unique_yards_for_item(kode_bahan, warna):
    barang_masuk_df = get_data_from_sheet('barang_masuk')
    barang_keluar_df = get_data_from_sheet('barang_keluar')
    
    yards_in = barang_masuk_df[(barang_masuk_df['kode_bahan'] == kode_bahan) & (barang_masuk_df['warna'] == warna)]['yard'].tolist()
    yards_out = barang_keluar_df[(barang_keluar_df['kode_bahan'] == kode_bahan) & (barang_keluar_df['warna'] == warna)]['yard'].tolist()
    
    unique_yards = sorted(list(set(yards_in + yards_out)))
    return unique_yards

def show_dashboard():
    st.title("Dashboard Bisnis 📈")
    st.markdown("---")

    master_df = get_master_barang()
    
    col_total_value, col_total_items = st.columns(2)
    if not master_df.empty:
        master_df['Stok Saat Ini'] = master_df.apply(lambda row: get_stock_balance(row['kode_bahan'], row['warna']), axis=1)
        master_df['harga'] = pd.to_numeric(master_df['harga'], errors='coerce').fillna(0)
        total_value = (master_df['Stok Saat Ini'] * master_df['harga']).sum()
        total_items = master_df['Stok Saat Ini'].sum()

        with col_total_value:
            st.metric("Total Nilai Stok Saat Ini", f"Rp {total_value:,.2f}")
        with col_total_items:
            st.metric("Total Barang di Gudang", f"{int(total_items)} Unit")
    else:
        st.info("Belum ada master barang untuk ditampilkan di dashboard.")

    st.markdown("---")
    st.header("Stok 10 Item Terendah")
    if not master_df.empty:
        low_stock_df = master_df.sort_values(by='Stok Saat Ini', ascending=True).head(10)
        low_stock_df['label'] = low_stock_df['nama_bahan'] + ' (' + low_stock_df['warna'] + ')'
        
        if not low_stock_df.empty:
            fig = px.bar(low_stock_df, 
                         x='label', 
                         y='Stok Saat Ini',
                         title='10 Item dengan Stok Terendah',
                         labels={'label': 'Nama Bahan', 'Stok Saat Ini': 'Jumlah Stok'},
                         color='Stok Saat Ini',
                         color_continuous_scale=px.colors.sequential.Sunset
                         )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Semua stok barang sudah habis.")
    else:
        st.info("Belum ada master barang untuk menampilkan grafik.")


def show_master_barang():
    st.title("Master Barang 📦")
    st.markdown("---")
    
    tab_add, tab_list = st.tabs(["➕ Tambah Barang Baru", "📝 Daftar Barang & Kelola"])
    
    with tab_add:
        with st.expander("Form Tambah Barang Baru", expanded=True):
            with st.form("add_item_form"):
                col1, col2 = st.columns(2)
                with col1:
                    kode_bahan = st.text_input("Kode Bahan").upper()
                    nama_supplier = st.text_input("Nama Supplier")
                    warna = st.text_input("Warna").lower()
                with col2:
                    nama_bahan = st.text_input("Nama Bahan")
                    rak = st.text_input("Rak")
                    harga = st.number_input("Harga", min_value=0.0)
                
                submitted = st.form_submit_button("💾 Simpan Barang")
                if submitted:
                    if add_master_item(kode_bahan, nama_supplier, nama_bahan, warna, rak, harga):
                        st.success(f"Barang **{nama_bahan}** dengan warna **{warna}** berhasil ditambahkan. ✅")
                        st.rerun()
                    else:
                        st.error("Kombinasi Kode Bahan dan Warna tersebut sudah ada. ❌")
    
    with tab_list:
        st.subheader("Daftar Barang")
        df = get_master_barang()
        if not df.empty:
            df_display = df.copy()
            df_display.rename(columns={
                'kode_bahan': 'Kode Bahan',
                'nama_supplier': 'Nama Supplier',
                'nama_bahan': 'Nama Bahan',
                'warna': 'Warna',
                'rak': 'Rak',
                'harga': 'Harga'
            }, inplace=True)
            st.dataframe(df_display, use_container_width=True, hide_index=True)
            
            st.markdown("---")
            with st.expander("Kelola Data Master"):
                # Buat dictionary untuk memetakan nama tampilan ke data asli
                item_options_map = {f"{row['kode_bahan']} ({row['warna']})": (row['kode_bahan'], row['warna']) for _, row in df.iterrows()}
                item_to_edit_str = st.selectbox("Pilih Kode Bahan (Warna)", list(item_options_map.keys()), key="select_edit_master")
                
                if item_to_edit_str:
                    # Ambil kode_bahan dan warna yang benar dari dictionary
                    selected_kode, selected_warna = item_options_map[item_to_edit_str]
                    
                    # Gunakan kode dan warna yang sudah dijamin benar untuk memfilter DataFrame
                    filtered_df = df[(df['kode_bahan'] == selected_kode) & (df['warna'] == selected_warna)]

                    if not filtered_df.empty:
                        selected_row = filtered_df.iloc[0]
                        
                        with st.form("edit_master_form"):
                            col1, col2 = st.columns(2)
                            with col1:
                                new_kode_bahan = st.text_input("Kode Bahan Baru", value=selected_row['kode_bahan']).upper()
                                new_nama_bahan = st.text_input("Nama Bahan", value=selected_row['nama_bahan'])
                                new_rak = st.text_input("Rak", value=selected_row['rak'])
                            with col2:
                                new_warna = st.text_input("Warna Baru", value=selected_row['warna']).lower()
                                new_nama_supplier = st.text_input("Nama Supplier", value=selected_row['nama_supplier'])
                                new_harga = st.number_input("Harga", value=selected_row['harga'], min_value=0.0)
                                
                            col_btn1, col_btn2 = st.columns(2)
                            with col_btn1:
                                if st.form_submit_button("Simpan Perubahan"):
                                    if update_master_item(selected_row['kode_bahan'], selected_row['warna'], new_kode_bahan, new_warna, new_nama_supplier, new_nama_bahan, new_rak, new_harga):
                                        st.success("Data berhasil diperbarui! ✅")
                                        st.rerun()
                                    else:
                                        st.error("Kombinasi Kode Bahan dan Warna baru sudah ada. Gagal menyimpan perubahan. ❌")
                            with col_btn2:
                                if st.form_submit_button("Hapus Barang"):
                                    delete_master_item(selected_row['kode_bahan'], selected_row['warna'])
                                    st.success("Data berhasil dihapus! 🗑️")
                                    st.rerun()
                    else:
                        st.warning("Data yang dipilih tidak ditemukan. Silakan refresh halaman atau pilih data lain.")
        else:
            st.info("Belum ada master barang.")

def show_input_masuk():
    st.title("Input Barang Masuk 📥")
    st.markdown("---")
    
    master_df = get_master_barang()
    if master_df.empty:
        st.warning("Belum ada master barang. Silakan tambahkan di menu Master Barang. ⚠️")
        return

    tab_add, tab_list = st.tabs(["➕ Input Barang Masuk Baru", "📝 Daftar Barang Masuk & Kelola"])

    with tab_add:
        with st.expander("Form Input Barang Masuk", expanded=True):
            with st.form("input_masuk_form"):
                col1, col2 = st.columns(2)
                with col1:
                    kode_bahan_options = master_df['kode_bahan'].unique().tolist()
                    selected_kode_bahan = st.selectbox("Pilih Kode Bahan", kode_bahan_options, key="in_kode_bahan")
                
                with col2:
                    filtered_colors = master_df[master_df['kode_bahan'] == selected_kode_bahan]['warna'].tolist()
                    selected_warna = st.selectbox("Warna", filtered_colors, key="in_warna")
                
                stok = st.number_input("Stok", min_value=1, key="in_stok")
                yard = st.number_input("Yard", min_value=0.0, key="in_yard")
                keterangan = st.text_area("Keterangan", key="in_keterangan")
                
                submitted = st.form_submit_button("💾 Simpan Barang Masuk")
                if submitted:
                    tanggal_waktu = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    add_barang_masuk(tanggal_waktu, selected_kode_bahan, selected_warna, stok, yard, keterangan)
                    st.success("Barang masuk berhasil dicatat. ✅")
                    st.rerun()
    
    with tab_list:
        st.subheader("Daftar Barang Masuk")
        df = get_barang_masuk()
        if not df.empty:
            df['tanggal_waktu'] = pd.to_datetime(df['tanggal_waktu']).dt.strftime('%Y-%m-%d %H:%M:%S')
            df.rename(columns={
                'tanggal_waktu': 'Tanggal & Waktu',
                'kode_bahan': 'Kode Bahan',
                'warna': 'Warna',
                'stok': 'Stok',
                'yard': 'Yard',
                'keterangan': 'Keterangan'
            }, inplace=True)
            st.dataframe(df.drop('id', axis=1), use_container_width=True, hide_index=True)

            st.markdown("---")
            with st.expander("Kelola Data Barang Masuk"):
                record_to_edit = st.selectbox("Pilih ID Data", df['id'].tolist(), key="select_edit_in")
                selected_row = df[df['id'] == record_to_edit].iloc[0]
                
                with st.form("edit_in_form"):
                    edit_tanggal_waktu = st.text_input("Tanggal & Waktu", value=selected_row['Tanggal & Waktu'])
                    kode_bahan_options = master_df['kode_bahan'].unique().tolist()
                    edit_kode_bahan = st.selectbox("Kode Bahan", kode_bahan_options, index=kode_bahan_options.index(selected_row['Kode Bahan']), key="edit_in_kode")
                    
                    filtered_colors_edit = master_df[master_df['kode_bahan'] == edit_kode_bahan]['warna'].tolist()
                    edit_warna = st.selectbox("Warna", filtered_colors_edit, index=filtered_colors_edit.index(selected_row['Warna']), key="edit_in_warna")
                    
                    edit_stok = st.number_input("Stok", value=selected_row['Stok'], min_value=1, key="edit_in_stok")
                    edit_yard = st.number_input("Yard", value=selected_row['Yard'], min_value=0.0, key="edit_in_yard")
                    edit_keterangan = st.text_area("Keterangan", value=selected_row['Keterangan'], key="edit_in_ket")

                    col_btn1, col_btn2 = st.columns(2)
                    with col_btn1:
                        if st.form_submit_button("Simpan Perubahan"):
                            update_barang_masuk(record_to_edit, edit_tanggal_waktu, edit_kode_bahan, edit_warna, edit_stok, edit_yard, edit_keterangan)
                            st.success("Data berhasil diperbarui! ✅")
                            st.rerun()
                    with col_btn2:
                        if st.form_submit_button("Hapus Data"):
                            delete_barang_masuk(record_to_edit)
                            st.success("Data berhasil dihapus! 🗑️")
                            st.rerun()
        else:
            st.info("Belum ada data barang masuk.")


def show_transaksi_keluar_invoice_page():
    st.title("Transaksi Keluar (Penjualan) & Invoice 🧾")
    st.markdown("---")
    
    tab_new_invoice, tab_history = st.tabs(["➕ Buat Transaksi & Invoice Baru", "📝 Riwayat Transaksi"])
    
    master_df = get_master_barang()
    if master_df.empty:
        st.warning("Belum ada master barang. Silakan tambahkan di menu Master Barang. ⚠️")
        return

    master_df['display_name'] = master_df['kode_bahan'] + ' - ' + master_df['nama_bahan'] + ' (' + master_df['warna'] + ')'
    item_options = master_df['display_name'].tolist()

    if 'cart_items' not in st.session_state:
        st.session_state['cart_items'] = []
    
    # Check for delete button clicks before rendering the form
    for i in range(len(st.session_state['cart_items'])):
        if st.session_state.get(f"delete_btn_{i}"):
            st.session_state['cart_items'].pop(i)
            st.session_state.pop(f"delete_btn_{i}") # Clean up session state
            st.rerun()

    with tab_new_invoice:
        st.subheader("Formulir Transaksi Penjualan")
        
        # Form untuk menambah item baru ke keranjang - di luar st.form utama
        with st.container(border=True):
            col_item_select, col_add_btn = st.columns([0.8, 0.2])
            with col_item_select:
                item_to_add_str = st.selectbox("Pilih Item yang Akan Dijual", item_options, key="item_add_select")
            with col_add_btn:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("➕ Tambah Item"):
                    selected_item_data = master_df[master_df['display_name'] == item_to_add_str].iloc[0]
                    new_item = {
                        "kode_bahan": selected_item_data['kode_bahan'],
                        "nama_bahan": selected_item_data['nama_bahan'],
                        "warna": selected_item_data['warna'],
                        "harga": selected_item_data['harga'],
                        "qty": 0,
                        "yard": 0,
                        "keterangan": ""
                    }
                    st.session_state['cart_items'].append(new_item)
                    st.rerun()

        # Start the single main form for transaction submission
        with st.form("new_transaction_form"):
            customer_name = st.text_input("Nama Pelanggan", help="Wajib diisi", key="customer_name")
            
            st.markdown("---")
            st.subheader("Keranjang Belanja 🛒")

            total_invoice = 0
            for i, item in enumerate(st.session_state['cart_items']):
                with st.container(border=True):
                    col_item_display, col_delete_btn = st.columns([0.9, 0.1])
                    
                    with col_item_display:
                        st.markdown(f"**Item {i+1}:** `{item['nama_bahan']} ({item['warna']})`")
                    # Move the st.button call outside of the form
                    # The previous check `st.session_state.get(f"delete_btn_{i}")` is now done before the form is rendered.
                    
                    stok_saat_ini = get_stock_balance(item['kode_bahan'], item['warna'])
                    
                    col_qty, col_yard = st.columns(2)
                    with col_qty:
                        min_val = 1 if stok_saat_ini > 0 else 0
                        max_val = int(stok_saat_ini)
                        st.session_state.cart_items[i]['qty'] = st.number_input(
                            "Jumlah",
                            min_value=min_val,
                            max_value=max_val,
                            value=int(st.session_state.cart_items[i].get('qty', 0)),
                            key=f"qty_{i}"
                        )
                    
                    with col_yard:
                        unique_yards = get_unique_yards_for_item(item['kode_bahan'], item['warna'])
                        yard_options = ["(Input Baru)"] + [str(y) for y in unique_yards]
                        
                        selected_yard_str = st.selectbox(
                            "Pilih Yard",
                            yard_options,
                            key=f"yard_select_{i}"
                        )

                        if selected_yard_str == "(Input Baru)":
                            st.session_state.cart_items[i]['yard'] = st.number_input(
                                "Input Yard Baru",
                                min_value=0.0,
                                key=f"yard_input_{i}"
                            )
                        else:
                            st.session_state.cart_items[i]['yard'] = float(selected_yard_str)
                    
                    st.session_state.cart_items[i]['keterangan'] = st.text_area(f"Keterangan (opsional)", value=st.session_state.cart_items[i].get('keterangan', ''), key=f"keterangan_{i}")
                    
                    current_item_total = st.session_state.cart_items[i]['qty'] * st.session_state.cart_items[i]['harga']
                    st.session_state.cart_items[i]['total'] = current_item_total
                    total_invoice += current_item_total
                    
                    st.markdown(f"**Harga Satuan:** Rp {st.session_state.cart_items[i]['harga']:,.2f}")
                    st.markdown(f"**Total Harga Item:** Rp {current_item_total:,.2f}")
                
            st.markdown(f"### **Total Keseluruhan:** **Rp {total_invoice:,.2f}**")
            
            submitted = st.form_submit_button("💾 Simpan Transaksi & Buat Invoice")
            if submitted:
                if not customer_name:
                    st.error("Nama Pelanggan wajib diisi.")
                elif not st.session_state['cart_items'] or all(item['qty'] == 0 for item in st.session_state['cart_items']):
                    st.error("Mohon tambahkan setidaknya satu item dengan jumlah lebih dari 0.")
                else:
                    new_invoice_number = generate_invoice_number()
                    success, message = add_barang_keluar_and_invoice(new_invoice_number, customer_name, st.session_state['cart_items'])
                    if success:
                        st.success(f"{message} Nomor Invoice: **{new_invoice_number}** ✅")
                        st.balloons()
                        st.session_state['cart_items'] = [] # Reset cart
                        st.rerun()
                    else:
                        st.error(message + " ❌")
                        
    with tab_history:
        st.subheader("Riwayat Transaksi")
        invoices_df = get_invoices()
        if not invoices_df.empty:
            invoices_df['tanggal_waktu'] = pd.to_datetime(invoices_df['tanggal_waktu']).dt.strftime('%Y-%m-%d %H:%M:%S')
            invoices_df.rename(columns={'invoice_number': 'No Invoice', 'tanggal_waktu': 'Tanggal & Waktu', 'customer_name': 'Nama Pelanggan'}, inplace=True)
            
            st.dataframe(invoices_df, use_container_width=True, hide_index=True)
            
            selected_invoice = st.selectbox("Pilih No Invoice untuk Dilihat/Unduh", invoices_df['No Invoice'].tolist())
            
            if st.button("Tampilkan & Unduh Invoice"):
                invoice_data = invoices_df[invoices_df['No Invoice'] == selected_invoice].iloc[0]
                invoice_items_df = get_invoice_items(selected_invoice)
                invoice_items_df['harga'] = pd.to_numeric(invoice_items_df['harga'], errors='coerce')
                invoice_items_df['total'] = pd.to_numeric(invoice_items_df['total'], errors='coerce')
                
                st.subheader(f"Detail Invoice: {selected_invoice}")
                st.write(f"**Nama Pelanggan:** {invoice_data['Nama Pelanggan']}")
                st.write(f"**Tanggal:** {invoice_data['Tanggal & Waktu']}")
                st.dataframe(invoice_items_df[['nama_bahan', 'qty', 'harga', 'total']].rename(columns={
                    'nama_bahan': 'Nama Bahan',
                    'qty': 'Qty',
                    'harga': 'Harga',
                    'total': 'Total'
                }), use_container_width=True, hide_index=True)

                pdf_file = generate_invoice_pdf(invoice_data, invoice_items_df)
                st.download_button(
                    label="Unduh Invoice PDF 📥",
                    data=pdf_file,
                    file_name=f"invoice_{selected_invoice}.pdf",
                    mime="application/pdf"
                )
        else:
            st.info("Belum ada riwayat transaksi.")

def show_monitoring_stok():
    st.title("Monitoring Stok 📊")
    st.markdown("---")
    
    st.subheader("Stok Saat Ini")
    master_df = get_master_barang()
    if not master_df.empty:
        master_df['Stok Saat Ini'] = master_df.apply(lambda row: get_stock_balance(row['kode_bahan'], row['warna']), axis=1)
        df_display = master_df[['kode_bahan', 'nama_bahan', 'warna', 'Stok Saat Ini']].copy()
        df_display.rename(columns={
            'kode_bahan': 'Kode Bahan',
            'nama_bahan': 'Nama Bahan',
            'warna': 'Warna',
            'Stok Saat Ini': 'Stok Saat Ini'
        }, inplace=True)
        st.dataframe(df_display, use_container_width=True, hide_index=True)
    else:
        st.warning("Belum ada master barang.")

    st.markdown("---")
    st.subheader("Rekam Jejak Stok (In & Out)")
    
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Tanggal Mulai", value=datetime.now().date())
    with col2:
        end_date = st.date_input("Tanggal Selesai", value=datetime.now().date())
        
    if st.button("Tampilkan Rekam Jejak"):
        records_df = get_in_out_records(start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d'))
        if not records_df.empty:
            records_df.rename(columns={
                'tanggal_waktu': 'Tanggal & Waktu',
                'kode_bahan': 'Kode Bahan',
                'warna': 'Warna',
                'qty': 'Jumlah',
                'type': 'Tipe',
                'keterangan': 'Keterangan'
            }, inplace=True)
            st.dataframe(records_df, use_container_width=True, hide_index=True)
        else:
            st.info("Tidak ada catatan stok masuk atau keluar pada rentang tanggal tersebut.")

def show_payroll_page():
    st.title("Sistem Penggajian Karyawan 💰")
    st.markdown("---")
    
    tab_master, tab_process, tab_history = st.tabs(["👥 Master Karyawan", "💸 Proses Penggajian", "📝 Riwayat Penggajian"])

    with tab_master:
        st.subheader("Data Master Karyawan")
        
        # Form to add new employee
        with st.expander("➕ Tambah Karyawan Baru", expanded=False):
            with st.form("add_employee_form"):
                col1, col2 = st.columns(2)
                with col1:
                    nama = st.text_input("Nama Karyawan")
                with col2:
                    bagian = st.text_input("Bagian")
                gaji = st.number_input("Gaji Pokok", min_value=0.0)
                
                submitted = st.form_submit_button("Tambah Karyawan")
                if submitted:
                    if nama and bagian and gaji > 0:
                        add_employee(nama, bagian, gaji)
                        st.success(f"Karyawan {nama} berhasil ditambahkan. ✅")
                        st.rerun()
                    else:
                        st.error("Semua field wajib diisi. ❌")

        st.markdown("---")
        st.subheader("Daftar Karyawan")
        employees_df_master = get_employees()
        if not employees_df_master.empty:
            employees_df_master.rename(columns={'id':'ID', 'nama_karyawan':'Nama', 'bagian':'Bagian', 'gaji_pokok':'Gaji Pokok'}, inplace=True)
            st.dataframe(employees_df_master, use_container_width=True, hide_index=True)

            # Form to edit/delete employee
            st.markdown("---")
            with st.expander("Kelola Data Karyawan"):
                selected_employee_id = st.selectbox("Pilih ID Karyawan", employees_df_master['ID'].tolist(), key='master_edit_select')
                
                selected_row = employees_df_master[employees_df_master['ID'] == selected_employee_id].iloc[0]
                
                with st.form("edit_employee_form"):
                    col1, col2 = st.columns(2)
                    with col1:
                        edit_nama = st.text_input("Nama", value=selected_row['Nama'])
                    with col2:
                        edit_bagian = st.text_input("Bagian", value=selected_row['Bagian'])
                    edit_gaji = st.number_input("Gaji Pokok", value=selected_row['Gaji Pokok'], min_value=0.0)
                    
                    col_btn1, col_btn2 = st.columns(2)
                    with col_btn1:
                        if st.form_submit_button("Simpan Perubahan"):
                            update_employee(selected_employee_id, edit_nama, edit_bagian, edit_gaji)
                            st.success("Data karyawan berhasil diperbarui! ✅")
                            st.rerun()
                    with col_btn2:
                        if st.form_submit_button("Hapus Karyawan"):
                            delete_employee(selected_employee_id)
                            st.success("Karyawan berhasil dihapus. 🗑️")
                            st.rerun()
        else:
            st.info("Belum ada data karyawan.")

    with tab_process:
        st.subheader("Proses Penggajian Bulanan")
        employees_df = get_employees()
        if employees_df.empty:
            st.warning("Tambahkan data karyawan terlebih dahulu di tab 'Master Karyawan'. ⚠️")
        else:
            st.markdown("Pilih karyawan yang akan diproses gajinya.")
            employee_options = employees_df.apply(lambda row: f"{row['id']} - {row['nama_karyawan']} ({row['bagian']})", axis=1).tolist()
            selected_employee_str = st.selectbox("Pilih Karyawan", employee_options)
            
            if selected_employee_str:
                employee_id = int(selected_employee_str.split(' - ')[0])
                selected_employee_data = employees_df[employees_df['id'] == employee_id].iloc[0]
                
                with st.form("payroll_form"):
                    st.write(f"**Nama:** {selected_employee_data['nama_karyawan']}")
                    st.write(f"**Bagian:** {selected_employee_data['bagian']}")

                    selected_date = st.date_input("Pilih Tanggal Gaji", value=datetime.now().date())
                    # Format the date to "Month Year" string for database storage
                    gaji_bulan = selected_date.strftime('%B %Y')
                    
                    st.markdown("### Pendapatan")
                    st.number_input("Gaji Pokok", value=selected_employee_data['gaji_pokok'], key="gaji_pokok_input", disabled=True)
                    lembur = st.number_input("Lembur", min_value=0.0)
                    lembur_minggu = st.number_input("Lembur Minggu", min_value=0.0)
                    uang_makan = st.number_input("Uang Makan", min_value=0.0)
                    
                    total_pendapatan = selected_employee_data['gaji_pokok'] + lembur + lembur_minggu + uang_makan
                    st.markdown(f"**Total Pendapatan (1):** **Rp {total_pendapatan:,.2f}**")
                    
                    st.markdown("### Potongan")
                    pot_absen_finger = st.number_input("Potongan Absen Finger", min_value=0.0)
                    ijin_hr = st.number_input("Ijin HR", min_value=0.0)
                    
                    total_setelah_potongan1 = total_pendapatan - pot_absen_finger - ijin_hr
                    st.markdown(f"**Total Setelah Potongan Absen (2):** **Rp {total_setelah_potongan1:,.2f}**")
                    
                    st.markdown("### Potongan Lain-lain")
                    simpanan_wajib = st.number_input("Simpanan Wajib", min_value=0.0)
                    potongan_koperasi = st.number_input("Potongan Koperasi", min_value=0.0)
                    kasbon = st.number_input("Kasbon", min_value=0.0)

                    gaji_akhir = total_setelah_potongan1 - simpanan_wajib - potongan_koperasi - kasbon
                    
                    st.markdown(f"### **TOTAL GAJI AKHIR:** **Rp {gaji_akhir:,.2f}**")
                    
                    keterangan = st.text_area("Keterangan", help="Opsional")
                    
                    submitted = st.form_submit_button("💾 Simpan Gaji")
                    if submitted:
                        add_payroll_record(employee_id, gaji_bulan, selected_employee_data['gaji_pokok'], lembur, lembur_minggu, uang_makan, pot_absen_finger, ijin_hr, simpanan_wajib, potongan_koperasi, kasbon, gaji_akhir, keterangan)
                        st.success(f"Penggajian untuk {selected_employee_data['nama_karyawan']} berhasil dicatat. ✅")
                        st.rerun()

    with tab_history:
        st.subheader("Riwayat Penggajian")
        
        # Filter for PDF download
        st.markdown("### Unduh Semua Slip Gaji (PDF)")
        payroll_months_df = get_data_from_sheet('payroll')
        if not payroll_months_df.empty:
            payroll_months = payroll_months_df['gaji_bulan'].unique().tolist()
            selected_month = st.selectbox("Pilih Bulan Gaji", payroll_months)
            
            if st.button(f"Unduh Slip Gaji {selected_month}"):
                payslip_data = get_payroll_records_by_month(selected_month)
                if not payslip_data.empty:
                    pdf_file = generate_payslips_pdf(payslip_data)
                    st.download_button(
                        label="Unduh PDF 📥",
                        data=pdf_file,
                        file_name=f"slip_gaji_{selected_month.replace(' ', '_')}.pdf",
                        mime="application/pdf"
                    )
                else:
                    st.error("Data penggajian tidak ditemukan untuk bulan tersebut. ❌")
        else:
            st.info("Tidak ada riwayat penggajian untuk diunduh.")

        st.markdown("---")
        st.subheader("Tabel Riwayat Penggajian")
        payroll_df = get_payroll_records()
        if not payroll_df.empty:
            # Mapping bulan dari Inggris ke Indonesia
            month_mapping = {
                'January': 'Januari', 'February': 'Februari', 'March': 'Maret',
                'April': 'April', 'May': 'Mei', 'June': 'Juni',
                'July': 'Juli', 'August': 'Agustus', 'September': 'September',
                'October': 'Oktober', 'November': 'November', 'December': 'Desember'
            }
            
            # Mengubah format tanggal_waktu
            payroll_df['tanggal_waktu'] = pd.to_datetime(payroll_df['tanggal_waktu'])
            payroll_df['tanggal_waktu'] = payroll_df['tanggal_waktu'].dt.strftime('%d %B %Y')
            
            # Mengganti nama bulan dari bahasa Inggris ke Indonesia
            for en, idn in month_mapping.items():
                payroll_df['tanggal_waktu'] = payroll_df['tanggal_waktu'].str.replace(en, idn)
            
            payroll_df.rename(columns={
                'tanggal_waktu':'Tanggal',
                'gaji_bulan': 'Periode Gaji',
                'nama_karyawan':'Nama Karyawan', 
                'gaji_akhir':'Gaji Bersih', 
                'keterangan':'Keterangan'
            }, inplace=True)
            st.dataframe(payroll_df.drop('id', axis=1), use_container_width=True, hide_index=True)
        else:
            st.info("Belum ada riwayat penggajian.")

# --- Login & Main App Logic ---
def login_page():
    st.title("Login Sistem Kontrol Stok")
    with st.form("login_form"):
        st.subheader("Silakan Masuk")
        username = st.text_input("Nama Pengguna")
        password = st.text_input("Kata Sandi", type="password")
        submitted = st.form_submit_button("Login")
        if submitted:
            if check_login(username, password):
                st.session_state['logged_in'] = True
                st.session_state['page'] = 'Dashboard'
                st.success("Berhasil Login! ✅")
                st.rerun()
            else:
                st.error("Nama pengguna atau kata sandi salah. ❌")

def main():
    if 'logged_in' not in st.session_state:
        st.session_state['logged_in'] = False
        st.session_state['page'] = 'Login'
    
    st.sidebar.title("PT. BERKAT KARYA ANUGERAH")
    st.sidebar.markdown("---")

    if st.session_state['logged_in']:
        if st.sidebar.button("Dashboard 📈", use_container_width=True):
            st.session_state['page'] = "Dashboard"
            st.rerun()
        if st.sidebar.button("Master Barang 📦", use_container_width=True):
            st.session_state['page'] = "Master Barang"
            st.rerun()
        if st.sidebar.button("Barang Masuk 📥", use_container_width=True):
            st.session_state['page'] = "Barang Masuk"
            st.rerun()
        if st.sidebar.button("Transaksi Keluar 🧾", use_container_width=True):
            st.session_state['page'] = "Transaksi Keluar"
            st.rerun()
        if st.sidebar.button("Monitoring Stok 📊", use_container_width=True):
            st.session_state['page'] = "Monitoring Stok"
            st.rerun()
        if st.sidebar.button("Penggajian 💰", use_container_width=True):
            st.session_state['page'] = "Penggajian"
            st.rerun()
        
        st.sidebar.markdown("---")
        if st.sidebar.button("Logout 🚪", use_container_width=True):
            st.session_state['logged_in'] = False
            st.session_state['page'] = 'Login'
            st.rerun()
        
        if st.session_state['page'] == "Dashboard":
            show_dashboard()
        elif st.session_state['page'] == "Master Barang":
            show_master_barang()
        elif st.session_state['page'] == "Barang Masuk":
            show_input_masuk()
        elif st.session_state['page'] == "Transaksi Keluar":
            show_transaksi_keluar_invoice_page()
        elif st.session_state['page'] == "Monitoring Stok":
            show_monitoring_stok()
        elif st.session_state['page'] == "Penggajian":
            show_payroll_page()
    else:
        login_page()

if __name__ == "__main__":
    init_app_state()
    main()
