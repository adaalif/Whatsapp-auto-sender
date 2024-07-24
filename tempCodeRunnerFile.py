import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import os
from Whatsapp_sender import MessageSender

class ExcelReaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Column Reader")
        self.root.geometry("1200x800")
        self.setup_styles()
        self.setup_widgets()
        self.data = []
        self.current_index = 0
        self.whatsapp_sender = MessageSender()

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.configure('TLabel', font=('Helvetica', 10))
        self.style.configure('TButton', font=('Helvetica', 10))
        self.style.configure('TEntry', font=('Helvetica', 10))

    def setup_widgets(self):
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
        self.label = ttk.Label(self.frame, text="Pilih file Excel:")
        self.label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.file_path_var = tk.StringVar()
        self.entry = ttk.Entry(self.frame, textvariable=self.file_path_var, width=40)
        self.entry.grid(row=0, column=1, padx=5, pady=5)
        self.browse_button = ttk.Button(self.frame, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=0, column=2, padx=5, pady=5)
        self.read_button = ttk.Button(self.frame, text="Read Columns", command=self.read_columns)
        self.read_button.grid(row=0, column=3, padx=5, pady=5)
        self.result_label = ttk.Label(self.frame, text="Nama-nama kolom:")
        self.result_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.text_frame = ttk.Frame(self.frame)
        self.text_frame.grid(row=2, column=0, padx=5, pady=5, sticky=tk.NSEW)
        self.result_text = tk.Text(self.text_frame, width=40, height=10, font=('Helvetica', 10), wrap=tk.WORD)
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar = ttk.Scrollbar(self.text_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.configure(yscrollcommand=self.scrollbar.set)
        self.template1_label = ttk.Label(self.frame, text="Template Pesan 1:")
        self.template1_label.grid(row=1, column=2, padx=5, pady=5, sticky=tk.W)
        self.template1_text = tk.Text(self.frame, width=40, height=10, font=('Helvetica', 10), wrap=tk.WORD)
        self.template1_text.grid(row=2, column=2, padx=5, pady=5, sticky=tk.NSEW)
        self.template2_label = ttk.Label(self.frame, text="Template Pesan 2:")
        self.template2_label.grid(row=1, column=3, padx=5, pady=5, sticky=tk.W)
        self.template2_text = tk.Text(self.frame, width=40, height=10, font=('Helvetica', 10), wrap=tk.WORD)
        self.template2_text.grid(row=2, column=3, padx=5, pady=5, sticky=tk.NSEW)
        self.view_message_button1 = ttk.Button(self.frame, text="View Message from Template 1", command=lambda: self.view_message(1))
        self.view_message_button1.grid(row=3, column=2, pady=10)
        self.view_message_button2 = ttk.Button(self.frame, text="View Message from Template 2", command=lambda: self.view_message(2))
        self.view_message_button2.grid(row=3, column=3, pady=10)
        self.send_button = ttk.Button(self.frame, text="Kirim", command=self.send_message)
        self.send_button.grid(row=4, column=0, columnspan=4, pady=10)
        self.message_label = ttk.Label(self.frame, text="Pesan yang Digenerate:")
        self.message_label.grid(row=5, column=0, padx=5, pady=5, sticky=tk.W)
        self.message_text = tk.Text(self.frame, width=80, height=10, font=('Helvetica', 10), wrap=tk.WORD)
        self.message_text.grid(row=6, column=0, columnspan=4, padx=5, pady=5, sticky=tk.NSEW)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.file_path_var.set(file_path)
            self.read_columns()

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
                if not df.empty:
                    self.data.extend(df.fillna('').astype(str).to_dict(orient='records'))
                self.result_text.insert(tk.END, "\n")
            self.current_index = 0
        except Exception as e:
            messagebox.showerror("Error", f"Error reading file: {e}")

    def view_message(self, template_number):
        if not self.data:
            messagebox.showwarning("Warning", "No data available!")
            return
        current_data = self.data[self.current_index]
        template_text = self.template1_text.get("1.0", tk.END).strip() if template_number == 1 else self.template2_text.get("1.0", tk.END).strip()
        message = template_text
        for key, value in current_data.items():
            message = message.replace(f"[{key}]", value)
        self.message_text.delete(1.0, tk.END)
        self.message_text.insert(tk.END, message)

    def send_message(self):
        if not self.data:
            messagebox.showwarning("Warning", "No data available to send messages!")
            return
        template_switch = 0
        for row_data in self.data:
            if template_switch % 2 == 0:
                template = self.template1_text.get("1.0", tk.END).strip()
            else:
                template = self.template2_text.get("1.0", tk.END).strip()
            message = template
            for key, value in row_data.items():
                message = message.replace(f"[{key}]", str(value))
            phone_number = row_data.get('No_HP', '')
            if message and phone_number:
                try:
                    self.whatsapp_sender.send_message(phone_number, message)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to send message to {phone_number}: {e}")
                    continue
            else:
                messagebox.showwarning("Warning", f"Phone number or message cannot be empty for {phone_number}.")
                continue
            template_switch += 1
        messagebox.showinfo("Success", "All messages sent successfully!")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelReaderApp(root)
    root.mainloop()