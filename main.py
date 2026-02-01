import warnings

warnings.filterwarnings("ignore", category=UserWarning)
import os, sys, subprocess, threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from docxtpl import DocxTemplate
from docxcompose.composer import Composer
from PIL import Image, ImageTk
from docx2pdf import convert
from docx import Document
from PyPDF2 import PdfMerger
import unicodedata
import matplotlib.pyplot as plt
import shutil

import customtkinter as ctk

BASE_OUT = "fams_output"
DOCX_OUT = os.path.join(BASE_OUT, "docx")
PDF_OUT = os.path.join(BASE_OUT, "pdf")
MERGED_DOCX_OUT = os.path.join(BASE_OUT, "merged_docx")
MERGED_PDF_OUT = os.path.join(BASE_OUT, "merged_pdf")
LOG_FILE = os.path.join(BASE_OUT, "fams_log.txt")
encodings = [
    "utf-8",
    "utf-8-sig",
    "latin1",
    "cp1252",
    "ascii",
    "utf-16",
    "utf-16-le",
    "utf-16-be",
    "cp850",
    "cp437",
    "mac_roman",
]
green_hex = "#228B22"
hover_green_hex = "#1F7E1F"


def sanitize_filename(name):
    name = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode("ascii")
    return "".join(c if c.isalnum() or c in " -_." else "_" for c in name)


def read_students(path, log_func=None):
    ext = os.path.splitext(path)[1].lower()
    if ext in [".csv", ".txt"]:
        for enc in encodings:
            try:
                df = pd.read_csv(path, dtype=str, encoding=enc)
                if log_func:
                    log_func(f"Successfully read file with encoding: {enc}")
                break
            except UnicodeDecodeError:
                if log_func:
                    log_func(f"Failed to read file with encoding: {enc}")
        else:
            err_msg = "Failed to decode file with all tried encodings."
            if log_func:
                log_func(err_msg)
            messagebox.showerror("Encoding Error", err_msg)
            raise ValueError(err_msg)
    elif ext in [".xls", ".xlsx"]:
        try:
            df = pd.read_excel(path, dtype=str)
            if log_func:
                log_func("Successfully read Excel file.")
        except Exception as e:
            err_msg = f"Failed to read Excel file: {e}"
            if log_func:
                log_func(err_msg)
            messagebox.showerror("File Read Error", err_msg)
            raise
    else:
        err_msg = "Unsupported file type. Use CSV or Excel."
        if log_func:
            log_func(err_msg)
        messagebox.showerror("File Type Error", err_msg)
        raise ValueError(err_msg)

    cols = {c.strip().lower(): c for c in df.columns}
    name_col, id_col = None, None
    for cand in ["name", "full name", "student name", "student"]:
        if cand in cols:
            name_col = cols[cand]
            break
    for cand in [
        "student_number",
        "student no",
        "student_no",
        "id",
        "studentid",
        "student number",
    ]:
        if cand in cols:
            id_col = cols[cand]
            break

    if name_col is None or id_col is None:
        if df.shape[1] >= 2:
            name_col, id_col = df.columns[0], df.columns[1]
            if log_func:
                log_func(
                    f"Using first two columns as Name and Student Number: {name_col}, {id_col}"
                )
        else:
            err_msg = "Could not detect name and student number columns."
            if log_func:
                log_func(err_msg)
            messagebox.showerror("Column Error", err_msg)
            raise ValueError(err_msg)

    students = []
    for _, row in df.iterrows():
        name = str(row.get(name_col, "")).strip()
        sid = str(row.get(id_col, "")).strip()
        if name and sid and sid.lower() != "nan":
            students.append({"name": name, "student_number": sid})

    if log_func:
        log_func(f"Loaded {len(students)} students successfully.")
    return students


class Worker(threading.Thread):
    def __init__(
        self,
        students,
        template_path,
        ui_callback,
        gen_pdf=False,
        merge_docx=False,
        merge_pdf=False,
    ):
        super().__init__()
        self.students = students
        self.template_path = template_path
        self.ui_callback = ui_callback
        self.gen_pdf = gen_pdf
        self.merge_docx = merge_docx
        self.merge_pdf = merge_pdf

    def run(self):
        global students
        total_steps = len(self.students)
        if self.gen_pdf:
            total_steps += 1
        if self.merge_docx:
            total_steps += 1
        if self.merge_pdf:
            total_steps += 1
        current_step = 0
        now = datetime.now().strftime("%B %d, %Y")
        docx_files, pdf_files = [], []

        for s in self.students:
            try:
                tpl = DocxTemplate(self.template_path)
                context = {
                    "name": s["name"],
                    "student_number": s["student_number"],
                    "date": now,
                }

                safe_name = sanitize_filename(s["name"])
                filename = f"{s['student_number']}_{safe_name}.docx"
                out_docx = os.path.join(DOCX_OUT, filename)

                os.makedirs(os.path.dirname(out_docx), exist_ok=True)

                tpl.render(context)
                tpl.save(out_docx)

                if not os.path.exists(out_docx):
                    raise FileNotFoundError(f"File not saved: {out_docx}")

                docx_files.append(out_docx)
                current_step += 1
                self.ui_callback(
                    progress=current_step / total_steps,
                    message=f"Generated DOCX: {filename}",
                )
            except Exception as e:
                current_step += 1
                self.ui_callback(
                    progress=current_step / total_steps,
                    message=f"Error for {s['name']}: {e}",
                )
        os.system("taskkill /f /im WINWORD.EXE >nul 2>&1")

        if self.gen_pdf:
            self.ui_callback(
                progress=current_step / total_steps,
                message="🕑 Starting PDF conversion...",
            )

            os.makedirs(PDF_OUT, exist_ok=True)

            total_files = len(docx_files)
            processed_files = 0

            for f in docx_files:
                try:
                    pdf_name = os.path.splitext(os.path.basename(f))[0] + ".pdf"
                    pdf_path = os.path.join(PDF_OUT, pdf_name)

                    # 🚨 EXE-SAFE: single file → single file
                    convert(os.path.abspath(f), os.path.abspath(pdf_path))

                    pdf_files.append(pdf_path)

                    self.ui_callback(message=f"✅ PDF created: {pdf_name}")

                except Exception as e:
                    self.ui_callback(
                        message=f"❌ PDF failed: {os.path.basename(f)} – {e}"
                    )

                finally:
                    processed_files += 1
                    self.ui_callback(
                        progress=current_step / total_steps
                        + processed_files / total_files / total_steps,
                        message=f"🕑 Converting ({processed_files}/{total_files})...",
                    )

            current_step += 1
            self.ui_callback(
                progress=current_step / total_steps,
                message="✅ PDF conversion completed",
            )

        if self.merge_docx and docx_files:
            try:
                master = Document(docx_files[0])
                composer = Composer(master)

                for f in docx_files[1:]:
                    composer.append(Document(f))

                merged_path = os.path.join(MERGED_DOCX_OUT, "MERGED_ALL.docx")
                composer.save(merged_path)

                current_step += 1
                self.ui_callback(
                    progress=current_step / total_steps,
                    message=f"✅ Merged DOCX saved: {merged_path}",
                )
            except Exception as e:
                current_step += 1
                self.ui_callback(
                    progress=current_step / total_steps,
                    message=f"❌ Error merging DOCX: {e}",
                )

        if self.merge_pdf and pdf_files:
            try:
                merger = PdfMerger()
                for f in pdf_files:
                    merger.append(f)
                merged_pdf_path = os.path.join(MERGED_PDF_OUT, "MERGED_ALL.pdf")
                merger.write(merged_pdf_path)
                merger.close()
                current_step += 1
                self.ui_callback(
                    progress=current_step / total_steps,
                    message=f"✅ Merged PDF saved: {merged_pdf_path}",
                )

            except Exception as e:
                current_step += 1
                self.ui_callback(
                    progress=current_step / total_steps,
                    message=f"❌ Error merging PDFs: {e}",
                )
        self.ui_callback(progress=1.0, message="✅ All tasks completed.", done=True)


class FAMSApp:
    def __init__(self, root):
        global topf
        self.root, self.students, self.template_path = root, [], None
        self.all_buttons = []
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.title("FAMS - Form Automation Management System")
        root.geometry("950x680")
        root.iconbitmap("assets/mbc.ico")
        root.configure(bg="#f0f4ff")
        style = ttk.Style(root)
        style.theme_use("clam")

        style.configure(
            "Amazing.Treeview",
            background="#ffffff",
            foreground="#333333",
            rowheight=28,
            fieldbackground="#ffffff",
            font=("Segoe UI", 10),
        )

        style.configure(
            "Amazing.Treeview.Heading",
            background="#2e4a9d",
            foreground="white",
            font=("Segoe UI", 11, "bold"),
            relief="raised",
        )

        style.map(
            "Amazing.Treeview",
            background=[("selected", "#c7d7ff")],
            foreground=[("selected", "#000000")],
        )

        style.map("Amazing.Treeview.Heading", background=[("active", "#1f3a8a")])

        style.configure(
            "Red.Horizontal.TProgressbar",
            troughcolor="#f0f0f0",
            background="#dd0f08",  # RED
            thickness=18,
        )

        style.configure(
            "TButton",
            font=("Segoe UI", 11, "bold"),
            foreground="white",
            background="#2e8b57",
            padding=6,
        )

        header_frame = tk.Frame(root, bg="#2e4a9d")
        header_frame.pack(fill="x")
        tk.Label(
            header_frame,
            text="Form Automation Management System ",
            font=("Segoe UI", 18, "bold"),
            bg="#2e4a9d",
            fg="white",
            pady=10,
        ).pack(side="left", padx=15)
        tk.Label(
            header_frame,
            text="Metro Business College",
            font=("Segoe UI", 10),
            bg="#2e4a9d",
            fg="white",
        ).place(x=20, y=60)
        try:
            img = Image.open("assets/mbc.png").resize((90, 90))
            self.logo_img = ImageTk.PhotoImage(img)
            tk.Label(header_frame, image=self.logo_img, bg="#2e4a9d").pack(
                side="right", padx=15
            )
        except Exception:
            tk.Label(
                header_frame,
                text="[Logo]",
                font=("Segoe UI", 12, "italic"),
                fg="white",
                bg="#2e4a9d",
            ).pack(side="right", padx=15)
        topf = tk.Frame(root, bg="#f0f4ff")
        topf.pack(fill="x", padx=12, pady=10)
        topf.grid_columnconfigure(3, weight=1)
        tk.Label(
            topf,
            text="① Upload Student File (CSV/Excel):",
            font=("Segoe UI", 11, "bold"),
            bg="#f0f4ff",
            fg="#2e4a9d",
        ).grid(row=0, column=0, sticky="w")
        self.lbl_file = tk.Label(
            topf, text="No file selected", fg="red", bg="#f0f4ff", font=("Segoe UI", 10)
        )
        self.lbl_file.grid(row=0, column=1, sticky="w", padx=8)
        # self.btn_file = ttk.Button(
        #     topf, text="Browse", command=self.browse_file, cursor="hand2"
        # )
        browse_img = Image.open("assets/browse.png").resize((20, 20))
        browse_imgtk = ImageTk.PhotoImage(browse_img)
        self.btn_file = ctk.CTkButton(
            topf,
            text="Browse",
            font=("Segoe UI", 13, "bold"),
            image=browse_imgtk,
            corner_radius=5,
            fg_color=green_hex,
            hover_color=hover_green_hex,
            cursor="hand2",
            width=110,
            height=30,
            command=self.browse_file,
        )
        self.btn_file.grid(row=0, column=2, padx=6)
        self.all_buttons.append(self.btn_file)

        tk.Label(
            topf,
            text="② Select DOCX Template:",
            font=("Segoe UI", 11, "bold"),
            bg="#f0f4ff",
            fg="#2e4a9d",
        ).grid(row=1, column=0, sticky="w", pady=6)
        self.lbl_template = tk.Label(
            topf,
            text="No template selected",
            fg="red",
            bg="#f0f4ff",
            font=("Segoe UI", 10),
        )
        self.lbl_template.grid(row=1, column=1, sticky="w")
        # self.btn_template = ttk.Button(
        #     topf, text="Browse", command=self.browse_template, cursor="hand2"
        # )
        self.btn_template = ctk.CTkButton(
            topf,
            text="Browse",
            font=("Segoe UI", 13, "bold"),
            image=browse_imgtk,
            corner_radius=5,
            fg_color=green_hex,
            hover_color=hover_green_hex,
            cursor="hand2",
            width=110,
            height=30,
            command=self.browse_template,
        )
        self.btn_template.grid(row=1, column=2, padx=6)
        self.all_buttons.append(self.btn_template)

        # ttk.Button(
        #     topf, text="🛈 Help", command=self.show_help, width=7, cursor="hand2"
        # ).grid(row=0, column=3, sticky="e", padx=10)
        help_img = Image.open("assets/help.png").resize((20, 20))
        help_imgtk = ImageTk.PhotoImage(help_img)
        ctk.CTkButton(
            topf,
            text="Help",
            # fg_color=green_hex,
            # hover_color=dark_green_hex,
            image=help_imgtk,
            cursor="hand2",
            font=("Segoe UI", 15, "bold"),
            text_color="#FFFFFF",
            width=10,
            command=self.show_help,
        ).grid(row=0, column=3, sticky="e", padx=10)

        midf = tk.Frame(root, bg="#e1e5f0")
        midf.pack(fill="both", expand=True, padx=12, pady=6)
        left = tk.LabelFrame(
            midf,
            text="Loaded Students",
            bg="#f0f4ff",
            fg="#2e4a9d",
            font=("Segoe UI", 11, "bold"),
        )
        left.pack(side="left", fill="both", expand=True, padx=(0, 6))
        scrollbar = ttk.Scrollbar(left, orient="vertical")
        scrollbar.pack(side="right", fill="y")
        self.tree = ttk.Treeview(
            left,
            columns=("sid", "name"),
            show="headings",
            height=14,
            style="Amazing.Treeview",
            yscrollcommand=scrollbar.set,
        )
        self.tree.tag_configure("oddrow", background="#f4f6fb")
        self.tree.tag_configure("evenrow", background="#ffffff")
        self.tree.heading("sid", text="Student Number")
        self.tree.heading("name", text="Student Name")
        self.tree.column("sid", width=160, anchor="center")
        self.tree.column("name", width=360, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=6, pady=4)

        scrollbar.config(command=self.tree.yview)
        self.lbl_count = tk.Label(
            left,
            text="0 students loaded",
            bg="#f0f4ff",
            fg="black",
            font=("Segoe UI", 10, "italic"),
        )
        self.lbl_count.pack(anchor="w", pady=4)
        right = tk.LabelFrame(
            midf,
            text="Actions & Log",
            bg="#f0f4ff",
            fg="#2e4a9d",
            font=("Segoe UI", 11, "bold"),
        )
        right.pack(side="right", fill="y")

        # ttk.Button(
        #     right,
        #     text="🚀 Generate Documents",
        #     command=self.start_generate,
        #     cursor="hand2",
        # ).pack(pady=(8, 6))
        gen_img = Image.open("assets/genrate.png").resize((30, 30))
        gen_imgtk = ImageTk.PhotoImage(gen_img)
        self.generate_btn = ctk.CTkButton(
            right,
            text="Generate Documents",
            image=gen_imgtk,
            # fg_color=green_hex,
            # hover_color=dark_green_hex,
            # text_color="black",
            cursor="hand2",
            font=("Segoe UI", 15, "bold"),
            text_color="#FFFFFF",
            width=10,
            command=self.start_generate,
        )
        self.generate_btn.pack(pady=(8, 6))
        self.all_buttons.append(self.generate_btn)

        self.pdf_var = tk.BooleanVar()
        self.merge_docx_var = tk.BooleanVar()
        self.merge_pdf_var = tk.BooleanVar()
        tk.Checkbutton(
            right,
            text="Generate PDF also",
            variable=self.pdf_var,
            bg="#f0f4ff",
            font=("Segoe UI", 10),
        ).pack()
        tk.Checkbutton(
            right,
            text="Merge all DOCX",
            variable=self.merge_docx_var,
            bg="#f0f4ff",
            font=("Segoe UI", 10),
        ).pack()
        tk.Checkbutton(
            right,
            text="Merge all PDFs",
            variable=self.merge_pdf_var,
            bg="#f0f4ff",
            font=("Segoe UI", 10),
        ).pack()
        # ttk.Button(
        #     right,
        #     text="📂 Open Output Folder",
        #     command=self.open_output,
        #     cursor="hand2",
        # ).pack(pady=6)
        openfolder_img = Image.open("assets/openfolder.png").resize((30, 30))
        openfoldertk = ImageTk.PhotoImage(openfolder_img)

        self.btn_open_output = ctk.CTkButton(
            right,
            text="Open Output Folder",
            image=openfoldertk,
            command=self.open_output,
            font=("Segoe UI", 12, "bold"),
            text_color="#FFFFFF",
            cursor="hand2",
        )
        self.btn_open_output.pack(pady=6)
        self.all_buttons.append(self.btn_open_output)

        btn_frame = tk.Frame(right, bg="#f0f4ff")
        btn_frame.pack(pady=6, fill="x")

        # ttk.Button(
        #     btn_frame, text="Download Logs", command=self.save_logs, cursor="hand2"
        # ).pack(side="left", expand=True, fill="x", padx=(3, 3))
        img_log = Image.open("assets/log.png").resize((20, 20))
        img_logtk = ImageTk.PhotoImage(img_log)
        self.btn_logs = ctk.CTkButton(
            btn_frame,
            text="Download Logs",
            font=("Segoe UI", 12, "bold"),
            fg_color="#7C0E0E",
            command=self.save_logs,
            cursor="hand2",
            image=img_logtk,
        )
        self.btn_logs.pack(side="left", expand=True, fill="x", padx=(3, 3))
        self.all_buttons.append(self.btn_logs)

        # ttk.Button(
        #     btn_frame,
        #     text="🧹 Clear Fields",
        #     command=self.clear_fields,
        #     cursor="hand2",

        # ).pack(
        #     side="left",
        #     expand=True,
        #     fill="x",
        #     padx=(3, 3),
        # )
        clear_img = Image.open("assets/clear.png").resize((20, 20))
        clear_imgtk = ImageTk.PhotoImage(clear_img)
        self.btn_clear = ctk.CTkButton(
            btn_frame,
            text="Clear Fields",
            image=clear_imgtk,
            command=self.clear_fields,
            fg_color="#EE0808",
            font=("Segoe UI", 12, "bold"),
            width=110,
        )
        self.btn_clear.pack(
            side="left",
            expand=True,
            fill="x",
            padx=(5, 5),
        )
        self.all_buttons.append(self.btn_clear)
        # tt
        # k.Button(right, text="❓ Help / User Guide", command=self.show_help).pack(
        #     pady=(6, 10)
        # )

        self.progress = ttk.Progressbar(
            right, length=260, mode="determinate", style="Red.Horizontal.TProgressbar"
        )
        self.progress.pack(pady=8)
        self.progress_label = tk.Label(
            right, text="0%", bg="#f0f4ff", font=("Segoe UI", 10, "bold")
        )
        self.progress_label.pack()
        tk.Label(
            right,
            text="Activity Log:",
            bg="#f0f4ff",
            fg="black",
            font=("Segoe UI", 10, "bold"),
        ).pack()
        self.log = tk.Text(
            right, width=36, height=12, state="disabled", font=("Consolas", 9)
        )
        self.log.pack(padx=6, pady=(0, 6))

    # def set_busy_ui(self):
    #     self.root.config(cursor="watch")
    #     for btn in self.all_buttons:
    #         btn.configure(state="disabled")
    #     self.root.update_idletasks()

    # def set_normal_ui(self):
    #     self.root.config(cursor="")
    #     for btn in self.all_buttons:
    #         btn.configure(state="normal")
    #     self.root.update_idletasks()

    def on_close(self):
        if messagebox.askokcancel("Exit", "Are you sure to exit FAMS?"):
            try:
                os._exit(0)  # HARD EXIT – guaranteed
            except:
                self.root.destroy()

    def browse_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xls;*.xlsx")]
        )
        if not path:
            return
        try:
            students = read_students(path)
            self.students = students
            self.lbl_file.configure(text=os.path.basename(path), fg="green")
            self.btn_file.configure(
                text="Browse",
            )

            self.check_browse = Image.open("assets/check.png").resize((30, 30))
            self.check_browse_tk = ImageTk.PhotoImage(self.check_browse)
            self.check_lbl_browse = tk.Label(
                topf, image=self.check_browse_tk, bg="#f0f4ff"
            )
            self.check_lbl_browse.grid(row=0, column=3, sticky="w")
            self.refresh_table()
            self.log_message(f"✅ Loaded {len(students)} students")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file: {e}")
            print(f"Failed to read file: {e}")

    def browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word Document", "*.docx")])
        if path and path.lower().endswith(".docx"):
            self.template_path = path
            self.lbl_template.configure(text=os.path.basename(path), fg="green")
            self.btn_template.configure(text="Browse")
            self.check_temp = Image.open("assets/check.png").resize((30, 30))
            self.check_temp_tk = ImageTk.PhotoImage(self.check_temp)
            self.check_lbl_temp = tk.Label(topf, image=self.check_temp_tk, bg="#f0f4ff")
            self.check_lbl_temp.grid(row=1, column=3, sticky="w")
            self.log_message(f"✅ Template selected: {os.path.basename(path)}")
        else:
            messagebox.showerror("Invalid", "Please select a DOCX template.")

    def refresh_table(self):
        for r in self.tree.get_children():
            self.tree.delete(r)

        for i, s in enumerate(self.students[:2000]):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.tree.insert(
                "", "end", values=(s["student_number"], s["name"]), tags=(tag,)
            )

        self.lbl_count.config(text=f"{len(self.students)} students loaded")

    def start_generate(self):

        if not self.students:
            messagebox.showwarning("No data", "Upload student data first.")
            return
        if not self.template_path:
            messagebox.showwarning("No template", "Select a DOCX template.")
            return
        for folder in [DOCX_OUT, PDF_OUT, MERGED_DOCX_OUT, MERGED_PDF_OUT]:
            os.makedirs(folder, exist_ok=True)
        self.progress["value"] = 0
        self.progress_label.config(text="0%")
        self.log_message("Starting...")

        Worker(
            self.students,
            self.template_path,
            self.worker_callback,
            self.pdf_var.get(),
            self.merge_docx_var.get(),
            self.merge_pdf_var.get(),
        ).start()

    def worker_callback(self, progress=0.0, message="", done=False):

        def update():
            self.progress["value"] = progress * 100
            self.progress_label.config(text=f"{int(progress*100)}%")
            if message:
                self.log_message(message)
            if done:

                messagebox.showinfo("Done", "All tasks completed.")

        self.root.after(0, update)

    def log_message(self, text):
        ts = datetime.now().strftime("%H:%M:%S")
        msg = f"[{ts}] {text}\n"
        self.log.config(state="normal")
        self.log.insert("end", msg)
        self.log.see("end")
        self.log.config(state="disabled")

    def save_logs(self):
        os.makedirs(BASE_OUT, exist_ok=True)
        logs = self.log.get("1.0", "end").strip()
        if not logs:
            messagebox.showinfo("Logs", "No logs to save.")
            return
        with open(LOG_FILE, "w", encoding="utf-8") as f:
            f.write(logs)
        messagebox.showinfo("Logs Saved", f"Logs saved to:\n{LOG_FILE}")

    def show_help(self):
        help_win = tk.Toplevel(self.root)
        help_win.title("FAMS – Help Guide")
        help_win.geometry("600x600")
        help_win.resizable(False, False)
        help_win.transient(self.root)
        help_win.iconbitmap("assets/mbc.ico")
        help_win.grab_set()
        logo_path = "assets/mbc.png"

        # Main frame with padding
        frame = tk.Frame(help_win, padx=15, pady=15)
        frame.pack(fill="both", expand=True)
        try:
            logo_img = Image.open(logo_path)
            logo_img = logo_img.resize(
                (100, 100), Image.LANCZOS
            )  # adjust size if needed
            self.help_logo = ImageTk.PhotoImage(logo_img)  # keep reference!

            logo_label = tk.Label(help_win, image=self.help_logo, borderwidth=0)
            logo_label.place(relx=1, x=-30, y=30, anchor="ne")  # top-right corner
        except Exception as e:
            print(f"Failed to load logo: {e}")

        # Canvas + Scrollbar for scrolling content
        canvas = tk.Canvas(frame, borderwidth=0)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas)

        scroll_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Helper function for section titles
        def section_title(text):
            return tk.Label(
                scroll_frame,
                text=text,
                font=("Segoe UI", 14, "bold"),
                fg="#2e4a9d",
                anchor="w",
                pady=8,
            )

        # Helper function for bullet points with icons
        def bullet_point(text, icon=None, fg="black"):
            frame = tk.Frame(scroll_frame, pady=2)
            frame.pack(anchor="w", fill="x")
            if icon:
                lbl_icon = tk.Label(
                    frame, text=icon, font=("Segoe UI Emoji", 12), fg=fg
                )
                lbl_icon.pack(side="left")
            lbl_text = tk.Label(
                frame, text=text, font=("Segoe UI", 11), justify="left", wraplength=480
            )
            lbl_text.pack(side="left", padx=6)
            return frame

        def ssexample(event=None):

            image_path = "assets/ss_example.png"

            if not os.path.exists(image_path):
                print(f"File not found: {image_path}")
                self.log_message(f"File not found: {image_path}")
                return

            # Option 1: Pillow
            try:
                from PIL import Image

                img = Image.open(image_path)
                img.show()
                return
            except Exception as e:
                print(f"PIL failed: {e}")

            # Option 2: Matplotlib
            try:
                import matplotlib.pyplot as plt
                from PIL import Image

                img = Image.open(image_path)
                plt.imshow(img)
                plt.axis("off")
                plt.show()
                return
            except Exception as e:
                print(f"Matplotlib failed: {e}")

            # Option 3: OS default viewer
            try:
                if sys.platform == "win32":
                    os.startfile(image_path)
                elif sys.platform == "darwin":
                    os.system(f"open {image_path}")
                else:
                    os.system(f"xdg-open {image_path}")
            except Exception as e:
                print(f"OS open failed: {e}")

        # Section 1: Upload Student File
        section_title("1. Upload Student File").pack(fill="x")
        bullet_point("Click 'Browse' under Upload Student File", icon="☑️").pack()
        bullet_point("Supported formats: CSV, XLS, XLSX").pack()
        bullet_point(
            "File must contain Name and Student Number columns\n(𝐞.𝐠 𝐬𝐭𝐮𝐝𝐞𝐧𝐭_𝐧𝐮𝐦𝐛𝐞𝐫 | 𝐧𝐚𝐦𝐞)"
        ).pack()
        example = tk.Label(
            scroll_frame,
            text="Screen Shot Example",
            font=("Segoe UI", 11, "underline"),
            fg="blue",
            cursor="hand2",
        )
        example.pack(anchor="w", pady=(0, 6))
        example.bind("<Button-1>", ssexample)

        # Section 2: Select DOCX Template
        section_title("2. Select DOCX Template").pack(fill="x")
        bullet_point("Choose a Word (.docx) template", icon="☑️").pack()
        bullet_point("Use placeholders like {{ name }}, {{ student_number }}").pack()

        # Section 2.1: Template Page Break & Formatting Rules
        section_title("📄 Template Page Break & Formatting").pack(fill="x")

        bullet_point(
            "Each student record generates ONE document or ONE page", icon="ℹ️"
        ).pack()

        bullet_point(
            "Use a manual Page Break (Ctrl + Enter) to separate pages", icon="📌"
        ).pack()

        bullet_point(
            "Do NOT add extra blank pages at the end of the template",
            icon="⚠️",
            fg="#a94442",
        ).pack()

        bullet_point(
            "Place all placeholders on the same page\n(e.g {{ name }}, {{ student_number }})"
        ).pack()

        bullet_point(
            "Avoid putting placeholders inside text boxes, shapes, or headers unless required",
            icon="🚫",
        ).pack()

        bullet_point(
            "If merging DOCX or PDF files, page breaks control final document layout",
            icon="🧩",
        ).pack()

        bullet_point(
            "Headers and footers are supported, but must contain valid placeholders only",
            icon="✔️",
        ).pack()

        # Section 3: Generate Documents
        section_title("3. Generate Documents").pack(fill="x")
        bullet_point("Click 🚀 Generate Documents", icon="▶️").pack()
        bullet_point("Optional:", fg="#555555").pack()
        bullet_point("Generate PDF", icon="✓").pack()
        bullet_point("Merge all DOCX", icon="✓").pack()
        bullet_point("Merge all PDFs", icon="✓").pack()

        # Section 4: Output Files
        section_title("4. Output Files").pack(fill="x")
        bullet_point(
            "Generated files are saved in the 'fams_output' folder", icon="📁"
        ).pack()

        # Clickable label to open output folder
        def open_output_folder(event=None):
            self.open_output()

        output_link = tk.Label(
            scroll_frame,
            text="Open Output Folder",
            font=("Segoe UI", 11, "underline"),
            fg="blue",
            cursor="hand2",
        )
        output_link.pack(anchor="w", pady=(0, 6))
        output_link.bind("<Button-1>", open_output_folder)

        # Section 5: Logs
        section_title("5. Logs").pack(fill="x")
        bullet_point("All actions appear in the Activity Log").pack()

        # Clickable label to download logs
        def save_logs(event=None):
            self.save_logs()

        logs_link = tk.Label(
            scroll_frame,
            text="Download Logs",
            font=("Segoe UI", 11, "underline"),
            fg="blue",
            cursor="hand2",
        )
        logs_link.pack(anchor="w", pady=(0, 10))
        logs_link.bind("<Button-1>", save_logs)

        # Tips section
        section_title("🛠 Tips").pack(fill="x")
        bullet_point("Do not open DOCX files while generating PDFs").pack()
        bullet_point("Make sure the template file is closed").pack()
        bullet_point("Large student lists may take time").pack()

        # Final encouragement
        final_lbl = tk.Label(
            scroll_frame,
            text="✅ You're good to go!",
            font=("Segoe UI", 12, "bold"),
            fg="#2e8b57",
            pady=10,
        )
        final_lbl.pack(anchor="w")

        # Close button
        ttk.Button(
            help_win, text="Close", command=help_win.destroy, cursor="hand2"
        ).pack(pady=10)

    def open_output(self):
        path = os.path.abspath(BASE_OUT)
        os.makedirs(path, exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])

    def clear_fields(self):
        self.students = []
        self.template_path = None
        self.lbl_file.config(text="No file selected", fg="red")
        self.lbl_template.config(text="No template selected", fg="red")

        if hasattr(self, "check_lbl_temp") and self.check_lbl_temp.winfo_ismapped():
            self.check_lbl_temp.grid_forget()

        if hasattr(self, "check_lbl_browse") and self.check_lbl_browse.winfo_ismapped():
            self.check_lbl_browse.grid_forget()

        self.btn_template.configure(text="Browse")

        for r in self.tree.get_children():
            self.tree.delete(r)
        self.lbl_count.config(text="0 students loaded")
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")
        self.progress["value"] = 0
        self.progress_label.config(text="0%")
        self.log_message("🧹 Cleared all fields.")


if __name__ == "__main__":
    splash = tk.Tk()
    splash.overrideredirect(True)
    splash.wm_attributes("-topmost", True)

    if sys.platform.startswith("win"):
        splash.configure(bg="#FF66C4")
        splash.wm_attributes("-transparentcolor", "#FF66C4")

    try:
        logo_img = Image.open("assets/splash.png").resize((600, 600))
        splash.logo = ImageTk.PhotoImage(logo_img)
        width, height = splash.logo.width(), splash.logo.height()
    except:
        splash.logo = None
        width, height = 200, 200

    screen_w = splash.winfo_screenwidth()
    screen_h = splash.winfo_screenheight()
    x = (screen_w // 2) - (width // 2)
    y = (screen_h // 2) - (height // 2)
    splash.geometry(f"{width}x{height}+{x}+{y}")

    label = tk.Label(
        splash, image=splash.logo, bg="#FF66C4", borderwidth=0, highlightthickness=0
    )
    label.pack()

    splash.update()
    splash.after(3000, splash.destroy)
    splash.mainloop()

    root = tk.Tk()
    app = FAMSApp(root)
    root.mainloop()
