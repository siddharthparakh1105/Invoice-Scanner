import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import pandas as pd
from extractor import process_invoice_file
import sys
import threading
from tkinter import font as tkFont

class TextRedirector(object):
    def __init__(self, widget):
        self.widget = widget

    def write(self, str):
        self.widget.configure(state='normal')
        self.widget.insert('end', str)
        self.widget.see('end')
        self.widget.configure(state='disabled')

    def flush(self):
        pass

class InvoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Extractor For Jain & Lunkad")
        self.root.geometry("1000x850")
        self.root.resizable(False, False)
        
        self.colors = {
            "bg_gradient_start": "#00416A",
            "bg_gradient_end": "#799F0C",
            "glass_bg": "#2c3e50",
            "card_1_bg": "#e91e63",
            "card_2_bg": "#ff9800",
            "card_3_bg": "#4caf50",
            "log_bg": "#34495e",
            "progress_bg": "#1a252f",
            "text": "#ffffff",
            "text_muted": "#bdc3c7",
            "button_fg": "#ffffff"
        }
        self.title_font = tkFont.Font(family="Segoe UI", size=28, weight="bold")
        self.card_title_font = tkFont.Font(family="Segoe UI", size=16, weight="bold")
        self.button_font = tkFont.Font(family="Segoe UI", size=12, weight="bold")
        self.status_font = tkFont.Font(family="Segoe UI", size=10)

        self.pdf_file_paths = []
        self.excel_file_path = ""
        self.create_widgets()

    def create_widgets(self):
        self.bg_canvas = tk.Canvas(self.root, highlightthickness=0)
        self.bg_canvas.pack(fill=tk.BOTH, expand=True)
        self.draw_gradient(self.bg_canvas, self.colors["bg_gradient_start"], self.colors["bg_gradient_end"])
        self.bg_canvas.bind("<Configure>", lambda e: self.draw_gradient(e.widget, self.colors["bg_gradient_start"], self.colors["bg_gradient_end"]))

        self.main_card = tk.Frame(self.bg_canvas, bg=self.colors["glass_bg"], padx=40, pady=30, relief='flat')
        self.main_card.place(relx=0.5, rely=0.45, anchor='center')

        title_label = tk.Label(self.main_card, text="Invoice Extractor For Jain & Lunkad", font=self.title_font, bg=self.colors["glass_bg"], fg=self.colors['text'])
        title_label.grid(row=0, column=0, columnspan=3, pady=(10, 20))

        api_key_frame = tk.Frame(self.main_card, bg=self.colors["glass_bg"])
        api_key_frame.grid(row=1, column=0, columnspan=3, pady=(0, 20), sticky='ew')

        api_key_label = tk.Label(api_key_frame, text="Gemini API Key:", font=self.button_font, bg=self.colors["glass_bg"], fg=self.colors['text'])
        api_key_label.pack(side=tk.LEFT, padx=(0, 10))

        self.api_key_entry = tk.Entry(api_key_frame, font=self.status_font, bg=self.colors['log_bg'], fg=self.colors['text'], relief='flat', width=70, show="*")
        self.api_key_entry.pack(side=tk.LEFT, fill='x', expand=True)

        card1 = self.create_step_card(self.main_card, "Step 1:\nUpload PDF Files", self.colors["card_1_bg"])
        card1.grid(row=2, column=0, sticky="ns", padx=15, pady=10)
        self.create_pdf_widgets(card1)

        card2 = self.create_step_card(self.main_card, "Step 2:\nSelect Existing Excel File", self.colors["card_2_bg"])
        card2.grid(row=2, column=1, sticky="ns", padx=15, pady=10)
        self.create_excel_widgets(card2)
        
        card3 = self.create_step_card(self.main_card, "Step 3:\nStart AI Extraction", self.colors["card_3_bg"])
        card3.grid(row=2, column=2, sticky="ns", padx=15, pady=10)
        self.create_action_widgets(card3)

        self.log_frame = self.create_step_card(self.main_card, "Real-Time Logs", self.colors["log_bg"])
        self.log_frame.grid(row=3, column=0, columnspan=3, sticky="ew", padx=15, pady=20)
        self.create_log_widgets(self.log_frame)

        self.progress_frame = tk.Frame(self.bg_canvas, bg=self.colors["progress_bg"], padx=40, pady=20, relief='flat')
        self.progress_frame.place(relx=0.5, rely=0.9, anchor='center', width=900)
        self.create_progress_widgets(self.progress_frame)

    def draw_gradient(self, canvas, color1, color2):
        canvas.delete("gradient")
        width, height = self.root.winfo_width(), self.root.winfo_height()
        (r1, g1, b1), (r2, g2, b2) = self.root.winfo_rgb(color1), self.root.winfo_rgb(color2)
        r_ratio, g_ratio, b_ratio = float(r2 - r1) / (width + height), float(g2 - g1) / (width + height), float(b2 - b1) / (width + height)
        for i in range(width + height):
            nr, ng, nb = int(r1 + (r_ratio * i)), int(g1 + (g_ratio * i)), int(b1 + (b_ratio * i))
            color = f"#{nr:04x}{ng:04x}{nb:04x}"
            canvas.create_line(0, i, i, 0, tags=("gradient",), fill=color)
        self.bg_canvas.lower("gradient")
        
    def create_step_card(self, parent, title, color):
        card = tk.Frame(parent, bg=color, padx=10, pady=20)
        card.config(highlightbackground=color, highlightthickness=10, relief='flat')
        title_label = tk.Label(card, text=title, font=self.card_title_font, bg=color, fg=self.colors['text'], justify='center')
        title_label.pack(pady=(0, 20))
        return card

    def create_pdf_widgets(self, parent):
        drop_zone = tk.Canvas(parent, bg=self.colors['glass_bg'], highlightthickness=0, width=200, height=150)
        drop_zone.pack(pady=10, padx=15)
        drop_zone.create_rectangle(5, 5, 195, 145, outline=self.colors['text_muted'], dash=(5, 3), width=2)
        drop_label = tk.Label(drop_zone, text="Drag & Drop Files Here\n(or click browse below)", font=self.status_font, bg=self.colors['glass_bg'], fg=self.colors['text_muted'])
        drop_label.place(relx=0.5, rely=0.5, anchor='center')
        
        self.select_button = tk.Button(parent, text="Browse Files...", font=self.button_font, fg=self.colors['button_fg'], bg=self.colors['card_1_bg'], activeforeground=self.colors['text'], activebackground=self.colors['card_1_bg'], relief='flat', command=self.select_pdfs, borderwidth=0)
        self.select_button.pack(pady=10, fill='x', padx=15)

    def create_excel_widgets(self, parent):
        self.excel_path_label = tk.Label(parent, text="No file selected", font=self.status_font, bg=self.colors['glass_bg'], fg=self.colors['text_muted'], wraplength=180, justify='center', height=8)
        self.excel_path_label.pack(pady=10, padx=15, fill='x')
        self.excel_button = tk.Button(parent, text="Browse...", font=self.button_font, fg=self.colors['button_fg'], bg=self.colors['card_2_bg'], activeforeground=self.colors['text'], activebackground=self.colors['card_2_bg'], relief='flat', command=self.select_excel, borderwidth=0)
        self.excel_button.pack(pady=10, fill='x', padx=15)

    def create_action_widgets(self, parent):
        self.extract_button = tk.Button(parent, text="Start Extraction", font=self.button_font, fg=self.colors['button_fg'], bg=self.colors['card_3_bg'], activeforeground=self.colors['text'], activebackground=self.colors['card_3_bg'], relief='flat', command=self.start_processing_thread, borderwidth=0, height=8)
        self.extract_button.pack(pady=10, fill='both', expand=True, padx=15)

    def create_log_widgets(self, parent):
        self.log_text = tk.Text(parent, height=6, bg=self.colors['glass_bg'], fg=self.colors['text_muted'], relief='flat', state='disabled', font=("Courier New", 9), borderwidth=0)
        self.log_text.pack(fill='both', expand=True, padx=15, pady=10)
        sys.stdout = TextRedirector(self.log_text)

    def create_progress_widgets(self, parent):
        self.progress_label = tk.Label(parent, text="Ready to process files", font=self.status_font, bg=self.colors["progress_bg"], fg=self.colors['text'])
        self.progress_label.pack(pady=(0, 10))
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Custom.Horizontal.TProgressbar",
                        background='#4caf50',
                        troughcolor=self.colors['glass_bg'],
                        borderwidth=0,
                        lightcolor='#4caf50',
                        darkcolor='#4caf50')
        
        self.progress_bar = ttk.Progressbar(parent, style="Custom.Horizontal.TProgressbar",
                                            length=800, mode='determinate')
        self.progress_bar.pack(pady=5)
        
        self.progress_percent = tk.Label(parent, text="0%", font=self.status_font, bg=self.colors["progress_bg"], fg=self.colors['text_muted'])
        self.progress_percent.pack(pady=(5, 0))

    def update_progress(self, current, total, message="Processing..."):
        if total > 0:
            progress_value = (current / total) * 100
            self.progress_bar['value'] = progress_value
            self.progress_percent.config(text=f"{int(progress_value)}%")
            self.progress_label.config(text=message)
        self.root.update_idletasks()

    def reset_progress(self):
        self.progress_bar['value'] = 0
        self.progress_percent.config(text="0%")
        self.progress_label.config(text="Ready to process files")

    def select_pdfs(self):
        paths = filedialog.askopenfilenames(title="Select PDF Invoices", filetypes=[("PDF files", "*.pdf")])
        if paths:
            self.pdf_file_paths = paths
            self.select_button.config(text=f"{len(paths)} Files Selected")
            self.reset_progress()

    def select_excel(self):
        path = filedialog.askopenfilename(title="Select Excel File to Append", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.excel_file_path = path
            self.excel_path_label.config(text=f"Append to:\n{os.path.basename(path)}")

    def start_processing_thread(self):
        thread = threading.Thread(target=self.process_files)
        thread.daemon = True
        thread.start()

    def process_files(self):
        api_key = self.api_key_entry.get()
        if not api_key:
            messagebox.showwarning("API Key Required", "Please enter your Google Gemini API key to proceed.")
            return
            
        if not self.pdf_file_paths:
            messagebox.showwarning("No Files Selected", "Please select one or more PDF files.")
            return
            
        self.extract_button.config(state=tk.DISABLED)
        all_extracted_rows, has_errors = [], False
        total_files = len(self.pdf_file_paths)
        
        for i, path in enumerate(self.pdf_file_paths):
            try:
                self.update_progress(i, total_files, f"Processing file {i+1} of {total_files}: {os.path.basename(path)}")
                print(f"[{i+1}/{total_files}] Processing: {os.path.basename(path)}...")
                rows = process_invoice_file(path, api_key)
                all_extracted_rows.extend(rows)
                self.update_progress(i+1, total_files, f"Completed {i+1} of {total_files} files")
            except Exception as e:
                has_errors = True
                print(f"ERROR: {e}")
                messagebox.showerror(f"Error processing {os.path.basename(path)}", str(e))
                break
                
        self.extract_button.config(state=tk.NORMAL)
        
        if has_errors:
            print("Processing stopped due to an error.")
            self.progress_label.config(text="Processing stopped due to error")
        elif all_extracted_rows:
            self.update_progress(total_files, total_files, "Saving to Excel...")
            print("AI processing complete. Saving to Excel...")
            self.save_to_excel(pd.DataFrame(all_extracted_rows))
            self.progress_label.config(text="Processing completed successfully!")
        else:
            print("Could not extract any data from the files.")
            self.progress_label.config(text="No data extracted from files")

    def save_to_excel(self, new_df):
        column_order = [
            'Invoice Date', 'Invoice No', 'Supplier Name', 'GSTIN/UIN', 'Consignor From Name', 'Consignor From GSTIN',
            'Item Name', 'HSN Code', 'QTY', 'Rate', 'Batch No', 'Exp Date', 'Amount', 'Narration'
        ]
        for col in column_order:
            if col not in new_df.columns:
                new_df[col] = "NA"

        final_df = new_df.reindex(columns=column_order).fillna("NA")
        output_path = self.excel_file_path
        if not output_path:
            output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Workbook", "*.xlsx")], title="Save Extracted Data As...")
            if not output_path:
                print("Save operation cancelled by user.")
                return
        try:
            if os.path.exists(output_path):
                print(f"Appending data to {os.path.basename(output_path)}...")
                with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    if 'Sheet1' in writer.book.sheetnames:
                        startrow = writer.book['Sheet1'].max_row
                        final_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=startrow)
                    else:
                        final_df.to_excel(writer, sheet_name='Sheet1', index=False)
            else:
                print(f"Creating new file: {os.path.basename(output_path)}...")
                final_df.to_excel(output_path, index=False, sheet_name='Sheet1')
            print("Export successful!")
            messagebox.showinfo("Success", f"Data exported to:\n{os.path.abspath(output_path)}")
        except Exception as e:
            print(f"Excel Export Error: {e}")
            messagebox.showerror("Export Error", f"An error occurred while saving to Excel:\n{e}\n\nCheck if the file is open.")

if __name__ == "__main__":
    try:
        from extractor import pytesseract
        pytesseract.get_tesseract_version()
    except Exception:
        messagebox.showerror("Tesseract Not Found", "Tesseract OCR is not found or is not configured correctly. Please check the installation instructions in README.md.")
        sys.exit(1)
        
    root = tk.Tk()
    app = InvoiceApp(root)
    root.mainloop()
