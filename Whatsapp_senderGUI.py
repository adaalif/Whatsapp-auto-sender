import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import os
from Whatsapp_sender import MessageSender

class ExcelReaderApp:
    def __init__(self, root):
        # Inisialisasi aplikasi dan pengaturan awal
        self.root = root
        self.root.title("Whatsapp Auto Messenger")
        self.root.geometry("1400x800")  # Perbesar ukuran jendela untuk menampung tutorial
        self.setup_styles()
        self.setup_widgets()
        self.data = []
        self.current_index = 0
        self.whatsapp_sender = MessageSender()

    def setup_styles(self):
        # Mengatur gaya untuk widget
        self.style = ttk.Style()
        self.style.configure('TLabel', font=('Helvetica', 10))
        self.style.configure('TButton', font=('Helvetica', 10))
        self.style.configure('TEntry', font=('Helvetica', 10))

    def setup_widgets(self):
        # Mengatur semua widget di dalam frame
        self.frame = ttk.Frame(self.root, padding="10 10 10 10")
        self.frame.pack(fill=tk.BOTH, expand=True)
        self.frame.grid_rowconfigure(1, weight=0)
        self.frame.grid_rowconfigure(2, weight=1)
        self.frame.grid_rowconfigure(3, weight=0)
        self.frame.grid_rowconfigure(4, weight=0)
        self.frame.grid_rowconfigure(5, weight=1)
        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_columnconfigure(1, weight=1)
        self.frame.grid_columnconfigure(2, weight=1)
        self.frame.grid_columnconfigure(3, weight=1)
        self.frame.grid_columnconfigure(4, weight=1)  # Kolom tambahan untuk tutorial
        
        # Variabel untuk menyimpan path file
        self.file_path_var = tk.StringVar()
        self.label = ttk.Label(self.frame, text="Pilih file Excel:")
        self.label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.entry = ttk.Entry(self.frame, textvariable=self.file_path_var, width=40)
        self.entry.grid(row=0, column=0, padx=5, pady=5, columnspan=2)  # Gabungkan kolom 1 dan 2 untuk entry
        self.browse_button = ttk.Button(self.frame, text="Select File", command=self.browse_file)
        self.browse_button.grid(row=0, column=2, padx=0, pady=5, sticky=tk.W)  # Geser tombol browser ke kolom 3

        # self.read_button = ttk.Button(self.frame, text="Read Columns", command=self.read_columns)
        # self.read_button.grid(row=0, column=3, padx=5, pady=5)
        
        self.result_label = ttk.Label(self.frame, text="Nama-nama kolom:")
        self.result_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.text_frame = ttk.Frame(self.frame)
        self.text_frame.grid(row=2, column=0, padx=5, pady=5, sticky=tk.NSEW)
        self.result_text = tk.Text(self.text_frame, width=40, height=10, font=('Helvetica', 10), wrap=tk.WORD)
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar = ttk.Scrollbar(self.text_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.configure(yscrollcommand=self.scrollbar.set, state=tk.DISABLED)
        
        self.template1_label = ttk.Label(self.frame, text="Template Pesan 1:")
        self.template1_label.grid(row=1, column=2, padx=5, pady=5, sticky=tk.W)
        self.template1_text = tk.Text(self.frame, width=40, height=10, font=('Helvetica', 10), wrap=tk.WORD)
        self.template1_text.grid(row=2, column=2, padx=5, pady=5, sticky=tk.NSEW)
        self.template2_label = ttk.Label(self.frame, text="Template Pesan 2:")
        self.template2_label.grid(row=1, column=3, padx=5, pady=5, sticky=tk.W)
        self.template2_text = tk.Text(self.frame, width=40, height=10, font=('Helvetica', 10), wrap=tk.WORD)
        self.template2_text.grid(row=2, column=3, padx=5, pady=5, sticky=tk.NSEW)
        
        self.phone_column_label = ttk.Label(self.frame, text="Kolom nomor handphone:")
        self.phone_column_label.grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
        self.phone_column_var = tk.StringVar()
        self.phone_column_combobox = ttk.Combobox(self.frame, textvariable=self.phone_column_var, state="readonly")
        self.phone_column_combobox.grid(row=4, column=1, padx=(0, 5), pady=5)  # Mengurangi padding kiri menjadi 0
        
        self.view_message_button1 = ttk.Button(self.frame, text="View Message from Template 1", command=lambda: self.view_message(1))
        self.view_message_button1.grid(row=4, column=2, pady=10)
        self.view_message_button2 = ttk.Button(self.frame, text="View Message from Template 2", command=lambda: self.view_message(2))
        self.view_message_button2.grid(row=4, column=3, pady=10)
        self.send_button = ttk.Button(self.frame, text="Kirim", command=self.send_message)
        self.send_button.grid(row=5, column=0, columnspan=4, pady=10)
        
        self.message_label = ttk.Label(self.frame, text="Pesan yang Digenerate:")
        self.message_label.grid(row=6, column=0, padx=5, pady=5, sticky=tk.W)
        self.message_text = tk.Text(self.frame, width=80, height=10, font=('Helvetica', 10), wrap=tk.WORD)
        self.message_text.grid(row=7, column=0, columnspan=4, padx=5, pady=5, sticky=tk.NSEW)
        self.message_text.configure(state=tk.DISABLED)

        # Input range of rows
        self.all_rows_var = tk.BooleanVar()
        self.all_rows_check = ttk.Checkbutton(self.frame, text="All Rows", variable=self.all_rows_var, command=self.toggle_row_inputs)
        self.all_rows_check.grid(row=8, column=0, padx=5, pady=5, sticky=tk.W)

        self.start_row_label = ttk.Label(self.frame, text="Start Row:")
        self.start_row_label.grid(row=9, column=0, padx=5, pady=5, sticky=tk.W)
        self.start_row_var = tk.StringVar()
        self.start_row_entry = ttk.Entry(self.frame, textvariable=self.start_row_var, width=20)
        self.start_row_entry.grid(row=9, column=1, padx=5, pady=5)
        
        self.end_row_label = ttk.Label(self.frame, text="End Row:")
        self.end_row_label.grid(row=9, column=2, padx=5, pady=5, sticky=tk.W)
        self.end_row_var = tk.StringVar()
        self.end_row_entry = ttk.Entry(self.frame, textvariable=self.end_row_var, width=20)
        self.end_row_entry.grid(row=9, column=3, padx=5, pady=5)

        # Panggil toggle_row_inputs awal untuk mengatur state yang benar
        self.toggle_row_inputs()

        # Bagian tutorial
        self.tutorial_label = ttk.Label(self.frame, text="Tutorial Template Pesan:")
        self.tutorial_label.grid(row=1, column=4, padx=5, pady=5, sticky=tk.W)
        self.tutorial_text = tk.Text(self.frame, width=40, font=('Helvetica', 10), wrap=tk.WORD)
        self.tutorial_text.grid(row=2, column=4, rowspan=5, padx=5, pady=5, sticky=tk.NSEW)
        self.tutorial_text.insert(tk.END, (
            "Tutorial Template Pesan:\n\n"
            "1. **Gunakan Placeholder**: Gunakan placeholder dalam template pesan seperti [Nama], [Alamat], dan [No_Hp].\n"
            "   - Contoh: 'Halo [Nama], alamat Anda adalah [Alamat] dan nomor telepon Anda adalah [No_Hp]'.\n\n"
            "2. **Masukkan Template Pesan**: Masukkan template pesan yang sesuai dalam 'Template Pesan 1' dan 'Template Pesan 2'.\n\n"
            "3. **Ganti Placeholder**: Program akan mengganti placeholder dengan data yang sesuai dari file Excel.\n\n"
            "4. **Preview dan Kirim**: Gunakan tombol 'View Message' untuk melihat pesan yang dihasilkan dari template, dan tombol 'Kirim' untuk mengirim pesan."
        ))

        # Mengatur agar text widget tidak bisa diedit
        self.tutorial_text.configure(state=tk.DISABLED)

    def toggle_row_inputs(self):
        # Mengatur state dari input baris berdasarkan checkbox "All Rows"
        if self.all_rows_var.get():
            self.start_row_entry.configure(state=tk.DISABLED)
            self.end_row_entry.configure(state=tk.DISABLED)
        else:
            self.start_row_entry.configure(state=tk.NORMAL)
            self.end_row_entry.configure(state=tk.NORMAL)

    def browse_file(self):
        # Membuka dialog untuk memilih file Excel
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.file_path_var.set(file_path)
            # Memuat file Excel untuk menemukan sheet yang tersedia
            df = pd.ExcelFile(file_path)
            sheet_names = df.sheet_names
            if len(sheet_names) > 1:
                selected_sheet = self.popup_sheet_selection(sheet_names)
                self.read_columns(selected_sheet)
            else:
                self.read_columns(sheet_names[0])

    def read_columns(self, sheet_name):
        # Membaca kolom dari sheet yang dipilih
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showwarning("Warning", "Please select a file first!")
            return
        if not os.path.exists(file_path):
            messagebox.showerror("Error", "File not found!")
            return
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            self.result_text.configure(state=tk.NORMAL)  # Ubah state menjadi NORMAL
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, f"Sheet: {sheet_name}\n")
            columns = list(df.columns)
            for idx, col in enumerate(columns, start=1):
                self.result_text.insert(tk.END, f"{idx}. {col}\n")
            self.result_text.configure(state=tk.DISABLED)  # Kembalikan state menjadi DISABLED
            if not df.empty:
                self.data = df.fillna('').astype(str).to_dict(orient='records')
            self.current_index = 0
            # Mengisi combobox kolom nomor telepon dengan nama kolom
            self.phone_column_combobox['values'] = columns
        except Exception as e:
            messagebox.showerror("Error", f"Error reading sheet {sheet_name}: {e}")

    def popup_sheet_selection(self, sheet_names):
        # Menampilkan popup untuk memilih sheet jika terdapat lebih dari satu sheet
        popup = tk.Toplevel(self.root)
        popup.title("Select a Sheet")
        popup.geometry("300x200")
        popup.transient(self.root)  # Set agar berada di atas jendela utama
        tk.Label(popup, text="Select a sheet:").pack(pady=10)

        var = tk.StringVar(popup)
        var.set(sheet_names[0])  # nilai default
        dropdown = ttk.Combobox(popup, textvariable=var, values=sheet_names, state="readonly")
        dropdown.pack()

        def on_select():
            selected_sheet = var.get()
            popup.destroy()  # Tutup popup setelah pemilihan
            return selected_sheet

        select_button = ttk.Button(popup, text="Select", command=on_select)
        select_button.pack(pady=20)

        popup.wait_window()  # Tunggu hingga popup ditutup
        return var.get()  # Kembalikan nama sheet yang dipilih

    def view_message(self, template_number):
        # Menampilkan pesan yang dihasilkan dari template
        if not self.data:
            messagebox.showwarning("Warning", "No data available!")
            return
        current_data = self.data[self.current_index]
        template_text = self.template1_text.get("1.0", tk.END).strip() if template_number == 1 else self.template2_text.get("1.0", tk.END).strip()
        message = template_text
        for key, value in current_data.items():
            message = message.replace(f"[{key}]", value)
        self.message_text.configure(state=tk.NORMAL)  # Ubah state menjadi NORMAL
        self.message_text.delete(1.0, tk.END)
        self.message_text.insert(tk.END, message)
        self.message_text.configure(state=tk.DISABLED)  # Kembalikan state menjadi DISABLED

    def send_message(self):
        # Mengirim pesan ke nomor telepon yang diambil dari kolom yang dipilih
        if not self.data:
            messagebox.showwarning("Warning", "No data available to send messages!")
            return
        phone_column = self.phone_column_var.get().strip()
        if not phone_column:
            messagebox.showwarning("Warning", "Tolong pilih kolom nomor handphone!")
            return
        
        # Verifikasi kolom nomor handphone
        first_phone_number = self.data[0].get(phone_column, '').strip()
        if first_phone_number.endswith('.0'):
            first_phone_number = first_phone_number[:-2]
        if not (first_phone_number.startswith('628') or first_phone_number.startswith('08') or first_phone_number.startswith('+628') or first_phone_number.startswith('8') or first_phone_number.startswith('62628') or first_phone_number.startswith('626208')):
            messagebox.showwarning("Warning", "Kolom yang dipilih bukan kolom nomor handphone yang valid! Tolong pilih kolom yang tepat.")
            return
        
        # Perbaiki kolom nomor handphone
        for row in self.data:
            phone_number = row.get(phone_column, '').strip()
            original_phone_number = phone_number  # Simpan nomor asli
            if phone_number.endswith('.0'):
                phone_number = phone_number[:-2]
            if phone_number.startswith('08'):
                phone_number = '628' + phone_number[2:]
            elif phone_number.startswith('8') and not phone_number.startswith('08'):
                phone_number = '628' + phone_number[1:]
            elif phone_number.startswith('6208'):
                phone_number = '628' + phone_number[4:]
            elif phone_number.startswith('+628'):
                phone_number = '628' + phone_number[4:]
            elif phone_number.startswith('62628'):
                phone_number = '628' + phone_number[5:]
            elif phone_number.startswith('626208'):
                phone_number = '628' + phone_number[6:]
            row[phone_column] = phone_number  # Simpan nomor yang sudah diperbaiki
        
        try:
            if self.all_rows_var.get():
                start_row = 0
                end_row = len(self.data) - 1
            else:
                start_row = int(self.start_row_var.get().strip()) - 1 if self.start_row_var.get().strip() else 0
                end_row = int(self.end_row_var.get().strip()) - 1 if self.end_row_var.get().strip() else len(self.data) - 1
        except ValueError:
            messagebox.showerror("Error", "Invalid row range! Please enter valid numbers.")
            return

        if start_row < 0 or end_row >= len(self.data) or start_row > end_row:
            messagebox.showerror("Error", "Invalid row range! Please enter a valid range within the data.")
            return

        template_switch = 0
        success = True  # Flag untuk melacak keberhasilan pengiriman pesan
        invalid_numbers = []  # List untuk melacak baris dengan nomor WhatsApp tidak valid
        empty_numbers = []  # List untuk melacak baris dengan nomor kosong atau tidak valid
        for index in range(start_row, end_row + 1):
            row_data = self.data[index]
            if template_switch % 2 == 0:
                template = self.template1_text.get("1.0", tk.END).strip()
            else:
                template = self.template2_text.get("1.0", tk.END).strip()
            message = template
            for key, value in row_data.items():
                message = message.replace(f"[{key}]", str(value))
            phone_number = row_data.get(phone_column, '').strip()
            if phone_number.endswith('.0'):
                phone_number = phone_number[:-2]
            if phone_number.startswith('08'):
                phone_number = '628' + phone_number[2:]
            elif phone_number.startswith('8') and not phone_number.startswith('08'):
                phone_number = '628' + phone_number[1:]
            elif phone_number.startswith('6208'):
                phone_number = '628' + phone_number[4:]
            elif phone_number.startswith('+628'):
                phone_number = '628' + phone_number[4:]
            elif phone_number.startswith('62628'):
                phone_number = '628' + phone_number[5:]
            elif phone_number.startswith('626208'):
                phone_number = '628' + phone_number[6:]
            if message and phone_number:
                try:

                    if not self.whatsapp_sender.send_message(phone_number, message):
                        invalid_numbers.append(index + 2)  # Menambahkan baris ke list jika nomor tidak valid di WhatsApp
                        success = False  # Set flag ke False jika ada kesalahan
                except Exception as e:
                    success = False  # Set flag ke False jika ada kesalahan
                    continue
            else:
                empty_numbers.append(index + 2)  # Menambahkan baris ke list jika nomor kosong atau tidak valid
                success = False  # Set flag ke False jika ada kesalahan
                continue
            template_switch += 1
        
        if success:
            messagebox.showinfo("Success", "All messages sent successfully!")
        else:
            error_message = "Some messages failed to send.\n"
            if invalid_numbers:
                error_message += f"Invalid WhatsApp numbers at rows: {', '.join(map(str, invalid_numbers))}\n"
            if empty_numbers:
                error_message += f"Empty or invalid phone numbers at rows: {', '.join(map(str, empty_numbers))}\n"
            messagebox.showerror("Error", error_message)

if __name__ == "__main__":
    # Inisialisasi dan jalankan aplikasi
    root = tk.Tk()
    app = ExcelReaderApp(root)
    root.mainloop()