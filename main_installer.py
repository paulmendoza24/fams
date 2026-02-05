import os, shutil, sys, threading, time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk, ImageDraw, ImageFont
import pythoncom
from win32com.shell import shell  # type: ignore

APP_NAME = "FAMS"
APP_TITLE = "Form Automation Management System"


def resource_path(path):
    try:
        base = sys._MEIPASS
    except Exception:
        base = os.path.abspath(".")
    return os.path.join(base, path)


SOURCE_DIR = resource_path("")
BANNER_PATH = resource_path(os.path.join("assets", "banner.png"))

DEFAULT_INSTALL_DIR = os.path.join(os.path.expanduser("~"), APP_NAME)
BANNER_PATH = os.path.join(SOURCE_DIR, "assets", "banner.png")


class InstallerWizard:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_NAME} Setup Wizard")
        self.root.geometry("700x500")
        self.root.resizable(False, False)
        self.root.iconbitmap(resource_path("assets/mbc.ico"))

        self.bg_main = "#f9f9fb"
        self.bg_side = "#2e4a9d"
        self.font_title = ("Segoe UI", 15, "bold")
        self.font_text = ("Segoe UI", 10)
        self.font_btn = ("Segoe UI", 10, "bold")

        self.container = tk.Frame(root, bg=self.bg_main)
        self.container.pack(fill="both", expand=True)

        self.sidebar = tk.Frame(self.container, bg=self.bg_side, width=160)
        self.sidebar.pack(side="left", fill="y")
        tk.Label(
            self.sidebar,
            text=APP_NAME,
            font=("Segoe UI", 18, "bold"),
            fg="white",
            bg=self.bg_side,
        ).pack(pady=40)
        tk.Label(
            self.sidebar,
            text="⚙️",
            font=("Segoe UI Emoji", 40),
            fg="white",
            bg=self.bg_side,
        ).pack(pady=10)

        self.page_frame = tk.Frame(self.container, bg=self.bg_main)
        self.page_frame.pack(side="right", fill="both", expand=True)

        self.banner_img = self.load_or_create_banner()

        self.current_page = 0
        self.pages = []
        self.install_dir = tk.StringVar(value=DEFAULT_INSTALL_DIR)
        self.create_shortcut = tk.BooleanVar(value=True)

        style = ttk.Style()
        style.configure("TButton", font=self.font_btn, padding=6)
        style.configure("TProgressbar", thickness=18)

        self.create_pages()
        self.show_page(0)

    def load_or_create_banner(self):

        if not os.path.exists(BANNER_PATH):
            os.makedirs(os.path.dirname(BANNER_PATH), exist_ok=True)
            img = Image.new("RGB", (540, 80), "#2e4a9d")

            for y in range(img.height):
                r = 46
                g = 74 + int((y / img.height) * 50)
                b = 157 + int((y / img.height) * 50)
                for x in range(img.width):
                    img.putpixel((x, y), (r, g, b))

            draw = ImageDraw.Draw(img)
            try:
                font = ImageFont.truetype("segoeui.ttf", 22)
            except:
                font = ImageFont.load_default()

            text = f"{APP_TITLE} Setup"
            bbox = draw.textbbox((0, 0), text, font=font)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
            draw.text(((540 - tw) / 2, (80 - th) / 2), text, font=font, fill="white")

            img.save(BANNER_PATH)

        img = Image.open(BANNER_PATH).resize((540, 80))
        return ImageTk.PhotoImage(img)

    def banner(self, parent):

        lbl = tk.Label(parent, image=self.banner_img, bg="white")
        lbl.image = self.banner_img
        lbl.pack(fill="x")

    def create_pages(self):

        page1 = tk.Frame(self.page_frame, bg=self.bg_main)
        self.banner(page1)
        tk.Label(
            page1,
            text=f"Welcome to the {APP_NAME} Setup Wizard",
            font=self.font_title,
            bg=self.bg_main,
            fg="#2e4a9d",
        ).pack(pady=30)
        tk.Label(
            page1,
            text=(
                "This wizard will install FAMS on your computer.\n\n"
                "It is recommended that you close all other applications before continuing.\n\n\n\n\n\n"
                "\t\t\t\t\tCreated by TEAM PAUL"
            ),
            font=self.font_text,
            bg=self.bg_main,
            justify="left",
        ).pack(padx=30)
        self.add_nav(page1, next_text="Next >", next_cmd=lambda: self.show_page(1))
        self.pages.append(page1)

        page2 = tk.Frame(self.page_frame, bg=self.bg_main)
        self.banner(page2)
        tk.Label(
            page2,
            text="Choose Destination Folder",
            font=self.font_title,
            bg=self.bg_main,
            fg="#2e4a9d",
        ).pack(pady=20)
        tk.Label(
            page2,
            text="Setup will install FAMS into the following folder:",
            font=self.font_text,
            bg=self.bg_main,
        ).pack()
        entry = tk.Entry(
            page2, textvariable=self.install_dir, width=55, font=("Consolas", 10)
        )
        entry.pack(pady=8)
        tk.Button(
            page2,
            text="Browse...",
            font=self.font_btn,
            bg="#2e8b57",
            fg="white",
            relief="flat",
            command=self.browse_folder,
        ).pack(pady=5)
        self.add_nav(
            page2,
            back_cmd=lambda: self.show_page(0),
            next_cmd=lambda: self.show_page(2),
        )
        self.pages.append(page2)

        page3 = tk.Frame(self.page_frame, bg=self.bg_main)
        self.banner(page3)
        tk.Label(
            page3,
            text="Installing...",
            font=self.font_title,
            bg=self.bg_main,
            fg="#2e4a9d",
        ).pack(pady=20)
        tk.Label(
            page3,
            text="Please wait while FAMS is being installed. Click Install.",
            font=self.font_text,
            bg=self.bg_main,
        ).pack()
        self.progress = ttk.Progressbar(page3, length=420, mode="determinate")
        self.progress.pack(pady=15)
        self.log_label = tk.Label(
            page3, text="", font=("Consolas", 9), bg=self.bg_main, fg="#2e4a9d"
        )
        self.log_label.pack()
        self.add_nav(
            page3,
            back_cmd=lambda: self.show_page(1),
            next_text="Install",
            next_cmd=self.start_install,
        )
        self.pages.append(page3)

        page4 = tk.Frame(self.page_frame, bg=self.bg_main)
        self.banner(page4)
        tk.Label(
            page4,
            text="Setup Complete",
            font=self.font_title,
            bg=self.bg_main,
            fg="green",
        ).pack(pady=30)
        tk.Label(
            page4,
            text=f"{APP_NAME} has been installed successfully!",
            font=self.font_text,
            bg=self.bg_main,
        ).pack(pady=5)
        tk.Checkbutton(
            page4,
            text=f"Create a desktop shortcut for {APP_TITLE}",
            variable=self.create_shortcut,
            font=self.font_text,
            bg=self.bg_main,
        ).pack(pady=20)
        self.add_nav(
            page4, back_cmd=None, next_text="Finish", next_cmd=self.finish_install
        )
        self.pages.append(page4)

    def add_nav(self, page, back_cmd=None, next_text="Next >", next_cmd=None):
        nav = tk.Frame(page, bg="#eeeeee")
        nav.pack(side="bottom", fill="x")
        if back_cmd:
            tk.Button(
                nav,
                text="< Back",
                font=self.font_btn,
                bg="#ddd",
                fg="black",
                width=10,
                relief="flat",
                command=back_cmd,
            ).pack(side="left", padx=20, pady=10)
        else:
            tk.Label(nav, width=10, bg="#eeeeee").pack(side="left", padx=20, pady=10)
        tk.Button(
            nav,
            text=next_text,
            font=self.font_btn,
            bg="#2e4a9d",
            fg="white",
            width=10,
            relief="flat",
            command=next_cmd,
        ).pack(side="right", padx=20, pady=10)

    def show_page(self, index):
        for page in self.pages:
            page.pack_forget()
        self.pages[index].pack(fill="both", expand=True)
        self.current_page = index

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.install_dir.set(os.path.join(folder, APP_NAME))

    def start_install(self):
        self.show_page(2)
        threading.Thread(target=self.install).start()

    def install(self):
        target_dir = self.install_dir.get()
        os.makedirs(target_dir, exist_ok=True)
        files = ["main.exe", "assets"]
        total = len(files)
        for i, f in enumerate(files, start=1):
            src = os.path.join(SOURCE_DIR, f)
            dst = os.path.join(target_dir, f)
            if os.path.isdir(src):
                if os.path.exists(dst):
                    shutil.rmtree(dst)
                shutil.copytree(src, dst)
            else:
                shutil.copy2(src, dst)
            time.sleep(0.7)
            self.progress["value"] = (i / total) * 100
            self.log_label.config(text=f"Installing {f}...")
            self.root.update_idletasks()
        self.show_page(3)

    def finish_install(self):
        if self.create_shortcut.get():
            self.create_desktop_shortcut()
        self.root.quit()

    def create_desktop_shortcut(self):
        try:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            shortcut_path = os.path.join(desktop, f"{APP_TITLE}.lnk")
            target = os.path.join(self.install_dir.get(), "main.exe")
            working_dir = self.install_dir.get()

            shell_link = pythoncom.CoCreateInstance(
                shell.CLSID_ShellLink,
                None,
                pythoncom.CLSCTX_INPROC_SERVER,
                shell.IID_IShellLink,
            )
            shell_link.SetPath(target)
            shell_link.SetDescription(APP_TITLE)
            shell_link.SetWorkingDirectory(working_dir)

            icon_path = os.path.join(working_dir, "assets", "app.ico")
            if os.path.exists(icon_path):
                shell_link.SetIconLocation(icon_path, 0)

            persist_file = shell_link.QueryInterface(pythoncom.IID_IPersistFile)
            persist_file.Save(shortcut_path, 0)

        except Exception as e:
            messagebox.showwarning("Shortcut Error", f"Could not create shortcut:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = InstallerWizard(root)
    root.mainloop()
