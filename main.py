# ===================================================================================
#  PDF LETTERHEAD MERGER - PRODUCTION SCRIPT
# ===================================================================================
#  - Author: Your Name
#  - Purpose: Merges a letterhead onto PDF files automatically in a watched folder.
#  - Features: Folder watching, manual merging, system tray functionality,
#              toast notifications, auto-start on login, and auto-updates from GitHub.
# ===================================================================================

import json
import os
import subprocess
import sys
import threading
import time
import winreg
from collections import defaultdict
from pathlib import Path
from tkinter import (
    BOTH,
    END,
    LEFT,
    RIGHT,
    StringVar,
    Text,
    Tk,
    Toplevel,
    filedialog,
    messagebox,
    ttk,
)

# --- 3rd Party Libraries ---
import pystray
import requests
from packaging.version import parse as parse_version
from PIL import Image
from pypdf import PdfReader, PdfWriter
from pystray import MenuItem as item
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

try:
    from ttkthemes import ThemedTk

    THEME_SUPPORT = True
except ImportError:
    THEME_SUPPORT = False


# ───── HELPER FUNCTION FOR PYINSTALLER (Finds bundled files) ─────
def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = getattr(sys, "_MEIPASS", None)
        if base_path is None:
            # Not a PyInstaller bundle, we're in development
            base_path = os.path.abspath(".")
    except Exception:
        # Not a PyInstaller bundle, we're in development
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# ───── APPLICATION CONSTANTS ─────
APP_NAME = "PDF Letterhead Merger"

# !!! IMPORTANT: UPDATE THIS FOR EACH NEW RELEASE !!!
# The auto-updater compares this version with the GitHub release tag.
# Example: For a release tagged 'v1.0.1', this should be "1.0.1".
__version__ = "1.0.0"

# !!! IMPORTANT: EDIT THIS LINE !!!
# Set this to your public GitHub repository in 'username/repository_name' format.
GITHUB_REPO = "YOUR_USERNAME/YOUR_REPONAME"

# --- Other Constants ---
CONFIG_FILE = Path.home() / f".{APP_NAME.lower().replace(' ', '_')}_config.json"
ICON_PATH = resource_path("icon.ico")
DEBOUNCE_SECONDS = 10


# ───── CORE APPLICATION LOGIC ─────


def load_config():
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r") as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                pass
    return {"letterhead_path": "", "watch_folder": ""}


def save_config(config):
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=4)


config = load_config()


def log_message(text_widget, message):
    if not text_widget:
        return
    text_widget.config(state="normal")
    text_widget.insert(END, f"{time.strftime('%H:%M:%S')} - {message}\n")
    text_widget.see(END)
    text_widget.config(state="disabled")


processed_files = set()


def merge_letterhead(
    app_instance,
    invoice_path: Path,
    letterhead_path: Path,
    output_path: Path = None,
    retries: int = 3,
):
    log_widget = app_instance.log_text
    try:
        if invoice_path.name.endswith(".merged.pdf"):
            return
        if str(invoice_path) in processed_files:
            return
        processed_files.add(str(invoice_path))

        invoice = PdfReader(str(invoice_path))
        letterhead = PdfReader(str(letterhead_path))
        if len(letterhead.pages) != 1:
            raise ValueError("Letterhead PDF must have exactly one page.")

        header_page = letterhead.pages[0]
        writer = PdfWriter()
        for page in invoice.pages:
            merged_page = page
            merged_page.merge_page(header_page)
            writer.add_page(merged_page)

        temp_path = (
            output_path if output_path else invoice_path.with_suffix(".merged.pdf")
        )

        for _ in range(retries):
            try:
                with open(temp_path, "wb") as f:
                    writer.write(f)
                if output_path is None:
                    os.replace(temp_path, invoice_path)
                    msg = f"Merged: {invoice_path.name}"
                    log_message(log_widget, f"[✓] {msg}")
                    app_instance.notify("Merge Complete", msg)
                else:
                    msg = f"Merged file saved:\n{temp_path}"
                    log_message(log_widget, f"[✓] Manual Merge Saved: {temp_path.name}")
                    messagebox.showinfo("Success", msg)
                return
            except PermissionError:
                time.sleep(1.5)
        raise Exception("Could not save file after multiple retries.")
    except Exception as e:
        error_msg = f"Merge failed for {invoice_path.name}: {e}"
        app_instance.notify("Merge Failed", str(error_msg))
        log_message(log_widget, f"[✗] {error_msg}")
        if output_path:
            messagebox.showerror("Merge Error", str(e))


def wait_until_file_ready(file_path: Path, timeout=10):
    last_size = -1
    for _ in range(timeout * 2):
        if not file_path.exists():
            time.sleep(0.5)
            continue
        try:
            current_size = file_path.stat().st_size
            if current_size == last_size and current_size > 0:
                return True
            last_size = current_size
        except FileNotFoundError:
            return False
        time.sleep(0.5)
    return False


class InvoiceHandler(FileSystemEventHandler):
    def __init__(self, app_instance, letterhead_path):
        self.app = app_instance
        self.letterhead_path = letterhead_path
        self.processed_files = defaultdict(float)
        self.start_time = time.time()

    def _should_process(self, file_path: Path) -> bool:
        now = time.time()
        try:
            if file_path.stat().st_mtime < self.start_time:
                return False
        except FileNotFoundError:
            return False
        file_key = str(file_path.resolve())
        if now - self.processed_files.get(file_key, 0) < DEBOUNCE_SECONDS:
            return False
        self.processed_files[file_key] = now
        return True

    def _handle_pdf(self, file_path: Path):
        if not wait_until_file_ready(file_path):
            log_message(
                self.app.log_text, f"[!] File never stabilized: {file_path.name}"
            )
            return
        merge_letterhead(self.app, file_path, self.letterhead_path)

    def _process_event(self, event):
        if (
            event.is_directory
            or not event.src_path.lower().endswith((".pdf"))
            or event.src_path.lower().endswith(".merged.pdf")
        ):
            return
        file_path = Path(event.src_path)
        if self._should_process(file_path):
            self._handle_pdf(file_path)

    def on_created(self, event):
        self._process_event(event)

    def on_modified(self, event):
        self._process_event(event)


def add_to_startup():
    if not getattr(sys, "frozen", False):
        return
    try:
        exe_path = sys.executable
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE
        ) as reg_key:
            winreg.SetValueEx(
                reg_key, APP_NAME, 0, winreg.REG_SZ, f'"{exe_path}" --start-minimized'
            )
    except Exception as e:
        print(f"[!] Failed to add to startup: {e}")


# ───── MAIN GUI APPLICATION CLASS ─────
class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.icon = None
        self.observer = None
        self.root.title(f"{APP_NAME} v{__version__}")
        self.root.geometry("550x550")
        self.root.minsize(500, 500)
        try:
            self.root.iconbitmap(ICON_PATH)
        except Exception:
            print(f"[!] Could not find or load window icon from path: {ICON_PATH}")
        self.letterhead_path = StringVar(value=config.get("letterhead_path", ""))
        self.watch_folder = StringVar(value=config.get("watch_folder", ""))
        self.status_text = StringVar()

        self.setup_styles()
        self.create_widgets()

        self.root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)
        add_to_startup()
        self.update_status()

        if self.letterhead_path.get() and self.watch_folder.get():
            self.toggle_watch()
        threading.Thread(target=self.check_for_updates, daemon=True).start()

    def setup_styles(self):
        style = ttk.Style(self.root)
        if THEME_SUPPORT:
            try:
                self.root.set_theme("arc")
            except Exception:
                style.theme_use("clam")
        else:
            style.theme_use("clam")
        style.configure("TButton", padding=6, font=("Segoe UI", 10))
        style.configure("TLabelframe", padding=10)
        style.configure("TLabelframe.Label", font=("Segoe UI", 11, "bold"))
        style.configure("Accent.TButton")

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="15 15 15 15")
        main_frame.pack(fill=BOTH, expand=True)

        config_frame = ttk.Labelframe(main_frame, text="1. Configuration")
        config_frame.pack(fill="x", pady=(0, 10))
        config_frame.columnconfigure(1, weight=1)
        ttk.Label(config_frame, text="Letterhead PDF:").grid(
            row=0, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Entry(
            config_frame, textvariable=self.letterhead_path, state="readonly"
        ).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(config_frame, text="Browse...", command=self.select_letterhead).grid(
            row=0, column=2, padx=5, pady=5
        )
        ttk.Label(config_frame, text="Watch Folder:").grid(
            row=1, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Entry(config_frame, textvariable=self.watch_folder, state="readonly").grid(
            row=1, column=1, padx=5, pady=5, sticky="ew"
        )
        ttk.Button(config_frame, text="Browse...", command=self.select_folder).grid(
            row=1, column=2, padx=5, pady=5
        )

        self.watch_btn = ttk.Button(
            main_frame,
            text="▶ Start Watching",
            command=self.toggle_watch,
            style="Accent.TButton",
        )
        self.watch_btn.pack(fill="x", ipady=5, pady=10)

        manual_frame = ttk.Labelframe(main_frame, text="2. Manual Operations")
        manual_frame.pack(fill="x", pady=10)
        btn_frame = ttk.Frame(manual_frame)
        btn_frame.pack(fill="x", expand=True, pady=(5, 5))
        btn_frame.columnconfigure((0, 1), weight=1)
        ttk.Button(
            btn_frame, text="📎 Manual Merge Single PDF", command=self.manual_merge
        ).grid(row=0, column=0, padx=5, sticky="ew")
        ttk.Button(
            btn_frame, text="📚 Batch Merge Multiple PDFs", command=self.batch_merge
        ).grid(row=0, column=1, padx=5, sticky="ew")

        log_frame = ttk.Labelframe(main_frame, text="Activity Log")
        log_frame.pack(fill=BOTH, expand=True, pady=10)
        self.log_text = Text(
            log_frame,
            height=8,
            wrap="word",
            state="disabled",
            font=("Consolas", 9),
            relief="solid",
            borderwidth=1,
        )
        log_scroll = ttk.Scrollbar(
            log_frame, orient="vertical", command=self.log_text.yview
        )
        self.log_text.config(yscrollcommand=log_scroll.set)
        log_scroll.pack(side=RIGHT, fill="y")
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)

        status_frame = ttk.Frame(self.root, relief="sunken", padding=(5, 2))
        status_frame.pack(side="bottom", fill="x")
        self.status_indicator = ttk.Label(status_frame, text="●", font=("", 14))
        self.status_indicator.pack(side=LEFT, padx=(0, 5))
        ttk.Label(status_frame, textvariable=self.status_text).pack(side=LEFT)

    def update_status(self):
        if self.observer and self.observer.is_alive():
            self.status_text.set(
                f"Active: Watching '{os.path.basename(self.watch_folder.get())}'"
            )
            self.status_indicator.config(foreground="#28a745")
        elif not self.letterhead_path.get() or not self.watch_folder.get():
            self.status_text.set("Awaiting Configuration")
            self.status_indicator.config(foreground="#ffc107")
        else:
            self.status_text.set("Stopped. Ready to watch.")
            self.status_indicator.config(foreground="#dc3545")

    def toggle_watch(self):
        if self.observer:
            self.observer.stop()
            self.observer.join()
            self.observer = None
            self.watch_btn.config(text="▶ Start Watching", style="Accent.TButton")
            log_message(self.log_text, "[~] Folder watching stopped.")
        else:
            if not self.letterhead_path.get() or not self.watch_folder.get():
                messagebox.showerror(
                    "Error",
                    "Please select both a Letterhead PDF and a Watch Folder first.",
                )
                return
            handler = InvoiceHandler(self, Path(self.letterhead_path.get()))
            self.observer = Observer()
            self.observer.schedule(
                handler, path=self.watch_folder.get(), recursive=False
            )
            self.observer.start()
            self.watch_btn.config(text="⏹ Stop Watching", style="TButton")
            log_message(
                self.log_text, f"[+] Started watching folder: {self.watch_folder.get()}"
            )
        self.update_status()

    def select_letterhead(self):
        path = filedialog.askopenfilename(
            title="Select Letterhead PDF", filetypes=[("PDF Files", "*.pdf")]
        )
        if path:
            self.letterhead_path.set(path)
            config["letterhead_path"] = path
            save_config(config)
            self.update_status()

    def select_folder(self):
        path = filedialog.askdirectory(title="Select Folder to Watch")
        if path:
            self.watch_folder.set(path)
            config["watch_folder"] = path
            save_config(config)
            if self.observer and self.observer.is_alive():
                self.toggle_watch()
                self.toggle_watch()
            self.update_status()

    def manual_merge(self):
        if not self.letterhead_path.get():
            messagebox.showerror("Error", "Please select a Letterhead PDF first.")
            return
        input_pdf = filedialog.askopenfilename(
            title="Select Invoice PDF", filetypes=[("PDF Files", "*.pdf")]
        )
        if not input_pdf:
            return
        output_pdf = filedialog.asksaveasfilename(
            title="Save Merged PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
        )
        if not output_pdf:
            return
        merge_letterhead(
            self, Path(input_pdf), Path(self.letterhead_path.get()), Path(output_pdf)
        )

    def batch_merge(self):
        if not self.letterhead_path.get():
            messagebox.showerror("Error", "Please select a Letterhead PDF first.")
            return
        files = filedialog.askopenfilenames(
            title="Select PDFs to Batch Merge", filetypes=[("PDF Files", "*.pdf")]
        )
        if not files:
            return
        output_folder = filedialog.askdirectory(title="Select Output Folder")
        if not output_folder:
            return
        for file_path in files:
            file = Path(file_path)
            out_path = Path(output_folder) / file.with_suffix(".merged.pdf").name
            merge_letterhead(self, file, Path(self.letterhead_path.get()), out_path)

    def get_tray_image(self):
        try:
            return Image.open(ICON_PATH)
        except Exception as e:
            print(f"[!] Could not load tray icon from {ICON_PATH}: {e}")
            return Image.new("RGB", (64, 64), "white")

    def minimize_to_tray(self):
        self.root.withdraw()
        if self.icon and self.icon.visible:
            return
        menu = pystray.Menu(
            item("Show App", self.restore_window, default=True),
            item("Exit", self.exit_app),
        )
        self.icon = pystray.Icon(
            APP_NAME.lower().replace(" ", "_"), self.get_tray_image(), APP_NAME, menu
        )
        self.icon.run_detached()

    def notify(self, title, message):
        if self.icon and self.icon.visible:
            self.icon.notify(message, title)

    def restore_window(self):
        if self.icon:
            self.icon.stop()
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def exit_app(self):
        if self.icon:
            self.icon.stop()
        if self.observer:
            self.observer.stop()
            self.observer.join()
        self.root.quit()

    def check_for_updates(self):
        if not getattr(sys, "frozen", False):
            print("[i] Skipping update check in dev mode.")
            return
        if GITHUB_REPO == "YOUR_USERNAME/YOUR_REPONAME":
            print("[!] GITHUB_REPO not set. Skipping update check.")
            return
        try:
            api_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            response = requests.get(api_url, timeout=5)
            response.raise_for_status()
            latest_release = response.json()
            latest_version_str = latest_release["tag_name"].lstrip("v")

            if parse_version(latest_version_str) > parse_version(__version__):
                for asset in latest_release.get("assets", []):
                    if asset["name"].endswith(".exe"):
                        self.root.after(
                            0,
                            self.prompt_for_update,
                            latest_version_str,
                            asset["browser_download_url"],
                        )
                        break
        except Exception as e:
            print(f"[!] Update check failed: {e}")

    def prompt_for_update(self, new_version, download_url):
        if messagebox.askyesno(
            "Update Available",
            f"A new version ({new_version}) is available. Would you like to download and install it?",
        ):
            self.start_update(download_url)

    def start_update(self, download_url):
        update_window = Toplevel(self.root)
        update_window.title("Updating...")
        update_window.geometry("300x80")
        update_window.resizable(False, False)
        update_window.transient(self.root)
        ttk.Label(update_window, text=f"Downloading {APP_NAME}...").pack(pady=10)
        progress = ttk.Progressbar(
            update_window, orient="horizontal", length=280, mode="determinate"
        )
        progress.pack(pady=5)

        def _download():
            try:
                current_dir = Path(sys.executable).parent
                new_exe_path = current_dir / "update.exe"
                updater_script_path = current_dir / "updater.bat"

                response = requests.get(download_url, stream=True)
                response.raise_for_status()
                total_size = int(response.headers.get("content-length", 0))
                progress["maximum"] = total_size

                with open(new_exe_path, "wb") as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                        progress["value"] += len(chunk)
                        update_window.update_idletasks()

                current_exe = sys.executable
                with open(updater_script_path, "w") as f:
                    f.write("@echo off\n")
                    f.write(f"echo Updating {APP_NAME}...\n")
                    f.write("timeout /t 2 /nobreak > nul\n")
                    f.write(f'move /y "{new_exe_path}" "{current_exe}"\n')
                    f.write("echo Relaunching...\n")
                    f.write(f'start "" "{current_exe}"\n')
                    f.write('(goto) 2>nul & del "%~f0"\n')

                subprocess.Popen(
                    [str(updater_script_path)],
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
                self.root.after(100, self.exit_app)

            except Exception as e:
                messagebox.showerror("Update Failed", f"Could not download update: {e}")
                update_window.destroy()

        threading.Thread(target=_download, daemon=True).start()


# ───── APPLICATION ENTRY POINT ─────
if __name__ == "__main__":
    if THEME_SUPPORT:
        root = ThemedTk(theme="arc")
    else:
        root = Tk()
        print("[!] For a better UI, run: pip install ttkthemes")

    app = PDFMergerApp(root)

    # Check if we should start minimized (e.g., from Windows startup)
    if "--start-minimized" in sys.argv:
        app.minimize_to_tray()

    root.mainloop()
