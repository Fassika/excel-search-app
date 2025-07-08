import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import pandas as pd
import subprocess
from functools import partial

class ExcelSearcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Search Tool")
        self.root.geometry("850x600")
        self.root.configure(bg="#f4f6f7")

        self.files = []
        self.search_results = []

        self.setup_styles()
        self.create_widgets()

    def setup_styles(self):
        style = ttk.Style()
        style.configure("TButton", font=("Segoe UI", 10), padding=6)
        style.configure("TLabel", font=("Segoe UI", 10), background="#f4f6f7")
        style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"))

    def create_widgets(self):
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        header = ttk.Label(frame, text="Excel Search Tool", style="Header.TLabel")
        header.pack(pady=(0, 10))

        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=10)

        self.select_btn = ttk.Button(button_frame, text="Select Excel Files", command=self.select_files)
        self.select_btn.grid(row=0, column=0, padx=5)

        self.clear_btn = ttk.Button(button_frame, text="Clear All", command=self.clear_all)
        self.clear_btn.grid(row=0, column=1, padx=5)

        self.search_entry = ttk.Entry(frame, width=80)
        self.search_entry.pack(pady=10)
        self.search_entry.insert(0, "Enter text, number, or keyword to search")

        self.search_btn = ttk.Button(frame, text="Search", command=self.search)
        self.search_btn.pack(pady=10)

        self.results_label = ttk.Label(frame, text="Search Results:")
        self.results_label.pack(pady=5)

        self.results_box = tk.Text(frame, height=20, wrap=tk.WORD, bg="#ffffff", font=("Segoe UI", 10))
        self.results_box.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self.results_box, command=self.results_box.yview)
        self.results_box.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def select_files(self):
        selected = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
        if selected:
            self.files.extend(selected)
            messagebox.showinfo("Files Selected", f"Total files selected: {len(self.files)}")

    def clear_all(self):
        self.files = []
        self.search_results = []
        self.results_box.delete("1.0", tk.END)
        self.search_entry.delete(0, tk.END)
        self.search_entry.insert(0, "Enter text, number, or keyword to search")

    def search(self):
        term = self.search_entry.get().strip()
        if not self.files:
            messagebox.showerror("No Files", "Please select at least one Excel file.")
            return
        if not term:
            messagebox.showerror("No Search Term", "Please enter a search term.")
            return

        self.results_box.delete("1.0", tk.END)
        self.search_results.clear()

        for file_path in self.files:
            try:
                xl = pd.ExcelFile(file_path)
                for sheet in xl.sheet_names:
                    df = xl.parse(sheet, dtype=str, keep_default_na=False)
                    for row_idx, row in df.iterrows():
                        for col_idx, cell in enumerate(row):
                            if term.lower() in str(cell).lower():
                                cell_addr = self.get_cell_address(row_idx, col_idx)
                                result = {
                                    "file": file_path,
                                    "sheet": sheet,
                                    "cell": cell_addr,
                                    "value": str(cell)
                                }
                                self.search_results.append(result)
            except Exception as e:
                self.results_box.insert(tk.END, f"Error reading {file_path}: {e}\n")

        if self.search_results:
            self.display_results()
        else:
            self.results_box.insert(tk.END, "No matches found.\n")

    def display_results(self):
        self.results_box.delete("1.0", tk.END)
        for idx, res in enumerate(self.search_results):
            display = (
                f"[{idx+1}] File: {os.path.basename(res['file'])}, "
                f"Sheet: {res['sheet']}, Cell: {res['cell']}, Value: {res['value']}\n"
            )
            self.results_box.window_create(tk.END, window=ttk.Label(self.results_box, text=display, wraplength=760))
            go_btn = ttk.Button(
                self.results_box,
                text="Go To",
                width=10,
                command=partial(self.go_to_cell, res["file"], res["sheet"], res["cell"])
            )
            self.results_box.window_create(tk.END, window=go_btn)
            self.results_box.insert(tk.END, "\n\n")

    def get_cell_address(self, row, col):
        col_letter = ""
        while col >= 0:
            col_letter = chr(col % 26 + ord('A')) + col_letter
            col = col // 26 - 1
        return f"{col_letter}{row + 2}"

    def go_to_cell(self, file_path, sheet_name, cell_ref):
        try:
            import win32com.client as win32
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = True
            wb = excel.Workbooks.Open(file_path)
            ws = wb.Sheets(sheet_name)
            ws.Activate()
            ws.Range(cell_ref).Select()
        except Exception:
            try:
                subprocess.Popen(['start', '', file_path], shell=True)
                messagebox.showinfo("Opened file", f"Please locate Sheet: {sheet_name}, Cell: {cell_ref}")
            except Exception as ex:
                messagebox.showerror("Failed to Open", f"Could not open file:\n{ex}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSearcherApp(root)
    root.mainloop()
