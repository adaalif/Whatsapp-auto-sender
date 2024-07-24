import wget
import zipfile
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import subprocess
import requests

def get_chrome_version_windows():
    try:
        # Query the registry for the version number
        registry_path = r"HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon"
        version = subprocess.check_output(
            f'reg query "{registry_path}" /v version', shell=True, text=True
        )
        # Extract the full version number
        version = version.strip().split()[-1]  # Get the full version string
        return version
    except subprocess.CalledProcessError as e:
        print("Failed to get Chrome version:", e)
        return None

# Get the full version of Chrome installed on the system
chrome_version = get_chrome_version_windows()

# Use the Chrome version to get the corresponding ChromeDriver version

# Continue with your script using `version_number`
download_url = f"https://chromedriver.storage.googleapis.com/{chrome_version}/chromedriver_win64.zip"

# Download the zip file using the URL built above
latest_driver_zip = wget.download(download_url, 'chromedriver.zip')

# Extract the zip file
with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
     zip_ref.extractall()  # You can specify the destination folder path here

# Delete the zip file downloaded above
os.remove(latest_driver_zip)# Define required column names
required_columns = {
    'nama_perusahaan': 'nama_perusahaan',
    'total_piutang': 'total_piutang',
    'No_HP': 'No_HP',
    'npp': 'npp'
}

# Define paths for ChromeDriver and user profile
driver_path = os.path.join(os.getcwd(), 'chromedriver.exe')
user_data_dir = os.path.join(os.getcwd(), 'profile')

# Initialize WebDriver with options
chrome_options = Options()
chrome_options.add_argument(f"user-data-dir={user_data_dir}")

service = Service(driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

driver.get('https://web.whatsapp.com')

def kirim_pesan(nomor, pesan):
    url = f'https://web.whatsapp.com/send?phone={nomor}&text={pesan}'
    driver.get(url)
    try:
        input_box = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
        )
        input_box.send_keys(Keys.ENTER)
        time.sleep(5)
        return True
    except Exception as e:
        print(f"Pesan gagal dikirim ke {nomor}. Kesalahan: {e}")
        return False

# Define required column names
required_columns = {
    'nama_perusahaan': 'nama_perusahaan',
    'total_piutang': 'total_piutang',
    'No_HP': 'No_HP',
    'npp': 'npp'
}

# Display column name update message
print(
    "Halo,\n\n"
    "Mohon untuk mengganti nama kolom tabel Excel saat ini dengan nama-nama kolom berikut:\n\n"
    "1. nama_perusahaan\n"
    "2. total_piutang1\n"
    "3. No_HP\n"
    "4. npp\n\n"
    "Periksa apakah nama kolom di file Excel Anda sesuai dengan daftar di atas.\n"
    "Jika tidak, silakan perbaiki nama kolom sesuai dengan daftar tersebut.\n\n"
    "Terima kasih atas perhatian dan kerjasamanya.\n\n"
)

input("Tekan Enter setelah Anda memeriksa dan memperbaiki nama kolom sesuai dengan daftar di atas...")

# Prompt for file path
while True:
    file_path = input("Please enter the data file path: ")
    if os.path.exists(file_path):
        data = pd.read_excel(file_path)
        # Check if the columns in the file contain the required names
        actual_columns = set(data.columns.to_list())
        required_column_names = set(required_columns.values())

        missing_columns = required_column_names - actual_columns

        if not missing_columns:
            print("Nama kolom sudah benar. Memproses data...")
            break
        else:
            print("Kolom berikut tidak ditemukan di file Excel:")
            print(", ".join(missing_columns))
            print("\nSilakan perbaiki nama kolom di file Excel sesuai dengan daftar di atas.")
            input("Tekan Enter setelah Anda memperbaiki nama kolom di file Excel...")
    else:
        print("File tidak ditemukan. Silakan coba lagi.")

# Initialize counters
total_sent = 0
total_failed = 0

# Iterate over each row and send alternating versions of the message
for index, row in data.iterrows():
    nomor = str(row['No_HP'])
    company_name = row['nama_perusahaan']
    nominal = row['total_piutang']
    npp = row['npp']
    
    message_versi_1 = (
        "Yth.\n"
        "Bapak/Ibu Pimpinan\n"
        f"{company_name}\n"
        "Terima kasih telah menjadi peserta BPJS Ketenagakerjaan, berdasarkan data kami sampai dengan JULI 2024, "
        "Bapak/Ibu belum melakukan pembayaran iuran dan denda BPJS Ketenagakerjaan.\n"
        f"Kami informasikan jumlah tagihan iuran BPJS Ketenagakerjaan {company_name} (NPP : {npp}) "
        f"periode tagihan s.d periode JULI 2024, sebesar {nominal},-.\n"
        "Bapak Ibu mohon informasinya apakah badan usaha ini masih berjalan, bila sudah tidak ada kegiatan mohon "
        "diinformasikan ke kami melalui wa ini untuk dapat kami proses selanjutnya.\n"
        "Hormat kami,\n"
        "BPJS Ketenagakerjaan"
    )

    message_versi_2 = (
        "Yth.\n"
        "Bapak/Ibu Pimpinan\n"
        f"{company_name}\n"
        "Terima kasih telah menjadi peserta BPJS Ketenagakerjaan. Bapak/Ibu Bersama ini disampaikan berdasarkan data kami "
        f"sampai dengan JULI 2024, Bapak/Ibu belum melakukan pembayaran iuran dan denda BPJS Ketenagakerjaan.\n"
        f"Kami informasikan jumlah tagihan iuran BPJS Ketenagakerjaan {company_name} (NPP : {npp}) "
        f"periode tagihan s.d periode JULI 2024, sebesar {nominal},-.\n"
        "Bapak Ibu mohon informasinya apakah badan usaha ini masih berjalan, bila sudah tidak ada kegiatan mohon "
        "diinformasikan ke kami melalui wa ini untuk dapat kami proses selanjutnya.\n"
        "Hormat kami,\n"
        "BPJS Ketenagakerjaan"
    )
    
    try:
        if index % 2 == 0:
            if kirim_pesan(nomor, message_versi_1):
                total_sent += 1
            else:
                total_failed += 1
        else:
            if kirim_pesan(nomor, message_versi_2):
                total_sent += 1
            else:
                total_failed += 1
        data.at[index, 'Status_terkirim'] = status

    except Exception as e:
        print(f"Terjadi kesalahan saat mengirim pesan ke {nomor}. Kesalahan: {e}")
        total_failed += 1
    finally:
        time.sleep(5)  # Wait 5 seconds before processing the next number

data.to_excel(file_path, index=False)
print(f"Jumlah pesan yang terkirim: {total_sent}")
print(f"Jumlah pesan yang tidak terkirim: {total_failed}")

driver.quit()