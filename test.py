import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import os
from selenium.webdriver.common.by import By

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import time

class ExcelReaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Column Reader")
        self.root.geometry("1200x800")  # Ukuran jendela yang lebih besar

        self.style = ttk.Style()
        self.style.configure('TLabel', font=('Helvetica', 10))
        self.style.configure('TButton', font=('Helvetica', 10))
        self.style.configure('TEntry', font=('Helvetica', 10))

        # Frame utama untuk menampung semua elemen
        self.frame = ttk.Frame(root, padding="10 10 10 10")
        self.frame.pack(fill=tk.BOTH, expand=True)

        # Baris pertama: Input file, Browse, dan Read Columns
        self.label = ttk.Label(self.frame, text="Pilih file Excel:")
        self.label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

        self.file_path_var = tk.StringVar()
        self.entry = ttk.Entry(self.frame, textvariable=self.file_path_var, width=40)
        self.entry.grid(row=0, column=1, padx=5, pady=5)

        self.browse_button = ttk.Button(self.frame, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=0, column=2, padx=5, pady=5)

        self.read_button = ttk.Button(self.frame, text="Read Columns", command=self.read_columns)
        self.read_button.grid(row=0, column=3, padx=5, pady=5)

        # Baris kedua: Menampilkan nama-nama kolom di kiri dan input template pesan di kanan
        self.result_label = ttk.Label(self.frame, text="Nama-nama kolom:")
        self.result_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)

        # Frame untuk menampilkan nama-nama kolom
        self.text_frame = ttk.Frame(self.frame)
        self.text_frame.grid(row=2, column=0, padx=5, pady=5, sticky=tk.NSEW)

        self.result_text = tk.Text(self.text_frame, width=40, height=10, font=('Helvetica', 10), wrap=tk.WORD)
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar = ttk.Scrollbar(self.text_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.configure(yscrollcommand=self.scrollbar.set)

        # Kolom untuk Template Pesan 1
        self.template1_label = ttk.Label(self.frame, text="Template Pesan 1:")
        self.template1_label.grid(row=1, column=2, padx=5, pady=5, sticky=tk.W)

        self.template1_text = tk.Text(self.frame, width=40, height=10, font=('Helvetica', 10), wrap=tk.WORD)
        self.template1_text.grid(row=2, column=2, padx=5, pady=5, sticky=tk.NSEW)

        # Kolom untuk Template Pesan 2
        self.template2_label = ttk.Label(self.frame, text="Template Pesan 2:")
        self.template2_label.grid(row=1, column=3, padx=5, pady=5, sticky=tk.W)

        self.template2_text = tk.Text(self.frame, width=40, height=10, font=('Helvetica', 10), wrap=tk.WORD)
        self.template2_text.grid(row=2, column=3, padx=5, pady=5, sticky=tk.NSEW)

        # Baris ketiga: Tombol Generate Message
        self.generate_button = ttk.Button(self.frame, text="Generate Message", command=self.generate_message)
        self.generate_button.grid(row=3, column=0, columnspan=4, pady=10)

        # Baris keempat: Tombol Kirim
        self.send_button = ttk.Button(self.frame, text="Kirim", command=self.send_message)
        self.send_button.grid(row=4, column=0, columnspan=4, pady=10)

        # Baris kelima: Judul dan Frame untuk Menampilkan Pesan yang Digenerate
        self.message_label = ttk.Label(self.frame, text="Pesan yang Digenerate:")
        self.message_label.grid(row=5, column=0, padx=5, pady=5, sticky=tk.W)

        self.message_text = tk.Text(self.frame, width=80, height=10, font=('Helvetica', 10), wrap=tk.WORD)
        self.message_text.grid(row=6, column=0, columnspan=4, padx=5, pady=5, sticky=tk.NSEW)

        # Data dari file Excel
        self.data = []
        self.current_index = 0
        self.generate_count = 0  # Counter untuk melacak jumlah generate

        # Path WebDriver
        self.driver_path = r"C:\chromedriver-win64\chromedriver.exe"
        self.user_data_dir = r'C:\chromedriver-win64\profile'

        # Atur grid row and column weights untuk responsif
        self.frame.grid_rowconfigure(1, weight=0)
        self.frame.grid_rowconfigure(2, weight=1)
        self.frame.grid_rowconfigure(3, weight=0)
        self.frame.grid_rowconfigure(4, weight=0)
        self.frame.grid_rowconfigure(5, weight=1)
        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_columnconfigure(1, weight=1)
        self.frame.grid_columnconfigure(2, weight=1)
        self.frame.grid_columnconfigure(3, weight=1)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.file_path_var.set(file_path)
            self.read_columns()  # Baca kolom data saat file dipilih

    def read_columns(self):
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showwarning("Warning", "Please select a file first!")
            return

        if not os.path.exists(file_path):
            messagebox.showerror("Error", "File not found!")
            return

        try:
            excel_data = pd.ExcelFile(file_path)
            sheet_names = excel_data.sheet_names
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, "Nama-nama kolom dalam setiap tabel di file Excel:\n\n")
            
            self.data = []
            for sheet in sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet)
                self.result_text.insert(tk.END, f"Sheet: {sheet}\n")
                columns = list(df.columns)
                for idx, col in enumerate(columns, start=1):
                    self.result_text.insert(tk.END, f"{idx}. {col}\n")
                
                # Simpan data dari setiap baris
                if not df.empty:
                    self.data.extend(df.fillna('').astype(str).to_dict(orient='records'))
                
                self.result_text.insert(tk.END, "\n")
            self.current_index = 0  # Reset index
            self.generate_count = 0  # Reset generate count
        except Exception as e:
            messagebox.showerror("Error", f"Error reading file: {e}")

    def generate_message(self):
        if not self.data:
            messagebox.showwarning("Warning", "No data available!")
            return

        if self.current_index >= len(self.data):
            messagebox.showinfo("Info", "No more data to generate messages.")
            return

        # Gantikan placeholder dengan data dari baris saat ini
        current_data = self.data[self.current_index]
        template = self.template1_text.get("1.0", tk.END).strip() if self.generate_count % 2 == 0 else self.template2_text.get("1.0", tk.END).strip()
        if not template:
            messagebox.showwarning("Warning", "Please enter both message templates!")
            return

        message = template
        for key, value in current_data.items():
            message = message.replace(f"[{key}]", value)
        
        # Tampilkan pesan di area yang ditentukan
        self.message_text.delete(1.0, tk.END)
        self.message_text.insert(tk.END, message)

        # Update index untuk data berikutnya
        self.current_index += 1
        self.generate_count += 1  # Update generate count

    def kirim_pesan(self, nomor, pesan):
        url = f'https://web.whatsapp.com/send?phone={nomor}&text={pesan}'
        driver = self.driver
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

    def send_message(self):
        if not self.data:
            messagebox.showwarning("Warning", "No data available!")
            return

        try:
            chrome_options = Options()
            chrome_options.add_argument(f"user-data-dir={self.user_data_dir}")
            service = Service(self.driver_path)
            self.driver = webdriver.Chrome(service=service, options=chrome_options)

            # Open WhatsApp Web
            self.driver.get('https://web.whatsapp.com/')
            time.sleep(15)  # Wait for manual login if needed

            for record in self.data:
                phone_number = record.get('Phone Number')  # Ganti dengan nama kolom nomor telepon yang sesuai
                if phone_number:
                    message = self.message_text.get("1.0", tk.END).strip()
                    if self.kirim_pesan(phone_number, message):
                        print(f"Pesan berhasil dikirim ke {phone_number}")
                    else:
                        print(f"Pesan gagal dikirim ke {phone_number}")

            self.driver.quit()
            messagebox.showinfo("Info", "Messages sent successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Error sending message: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelReaderApp(root)
    root.mainloop()