import os
import sys
import csv
import time
import threading
import subprocess
import logging
import traceback
import concurrent.futures
from urllib.parse import urlparse
from datetime import datetime

# Attempt to import required modules; bootstrap will handle if not present.
try:
    import pdfkit
    from PyPDF2 import PdfMerger
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError:
    pass  # Bootstrap code below will install these

# =============================================================================
# Bootstrap Virtual Environment & Dependency Installation
# =============================================================================
def bootstrap_venv():
    """
    If not running inside a virtual environment, create one in a folder 'venv',
    install dependencies, and relaunch the script.
    """
    if sys.prefix != sys.base_prefix:
        return  # Already in a venv

    venv_path = os.path.join(os.getcwd(), "venv")
    if not os.path.exists(venv_path):
        print("Creating virtual environment...")
        subprocess.check_call([sys.executable, "-m", "venv", "venv"])
        print("Virtual environment created.")

    # Determine pip and python executables in the venv.
    if os.name == "nt":
        pip_executable = os.path.join(venv_path, "Scripts", "pip.exe")
        python_executable = os.path.join(venv_path, "Scripts", "python.exe")
    else:
        pip_executable = os.path.join(venv_path, "bin", "pip")
        python_executable = os.path.join(venv_path, "bin", "python")

    # Install required dependencies.
    deps = ["pdfkit", "PyPDF2"]
    print("Installing dependencies...")
    subprocess.check_call([pip_executable, "install"] + deps)
    print("Dependencies installed.")

    # Relaunch the script in the virtual environment.
    print("Relaunching script in virtual environment...")
    subprocess.check_call([python_executable, __file__])
    sys.exit()

bootstrap_venv()

# =============================================================================
# Logging Setup
# =============================================================================
logging.basicConfig(
    filename="conversion.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# Global cancel flag.
cancel_event = threading.Event()

# =============================================================================
# Simple Tooltip Class for Tkinter Widgets
# =============================================================================
class ToolTip:
    def __init__(self, widget, text='widget info'):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        self.id = None
        widget.bind("<Enter>", self.enter)
        widget.bind("<Leave>", self.leave)

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(500, self.showtip)

    def unschedule(self):
        id_ = self.id
        self.id = None
        if id_:
            self.widget.after_cancel(id_)

    def showtip(self, event=None):
        if self.tipwindow or not self.text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 1
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

# =============================================================================
# Main Application Class
# =============================================================================
class WebPage2PDFBundle:
    def __init__(self, root):
        self.root = root
        self.root.title("WebPage2PDF Bundle")
        self.start_time = None
        self.advanced_visible = tk.BooleanVar(value=True)
        self.setup_gui()

    def setup_gui(self):
        # --------------------
        # Row 0: CSV File Selection
        tk.Label(self.root, text="CSV File:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.csv_entry = tk.Entry(self.root, width=50)
        self.csv_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(self.root, text="Browse", command=self.browse_csv).grid(row=0, column=2, padx=5, pady=5)

        # --------------------
        # Row 1: CSV Options with Help Icons
        # CSV header row
        self.csv_header_var = tk.BooleanVar(value=False)
        self.header_check = tk.Checkbutton(self.root, text="CSV has header row", variable=self.csv_header_var)
        self.header_check.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        header_help = tk.Label(self.root, text="?", fg="blue", cursor="question_arrow")
        header_help.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        ToolTip(header_help, "If your CSV file includes a header row (with column names), check this box to skip it during processing.")

        # URL Column Index
        tk.Label(self.root, text="URL Column Index (0-based):").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.csv_column_entry = tk.Entry(self.root, width=5)
        self.csv_column_entry.grid(row=1, column=3, padx=5, pady=5, sticky="w")
        self.csv_column_entry.insert(0, "0")
        col_help = tk.Label(self.root, text="?", fg="blue", cursor="question_arrow")
        col_help.grid(row=1, column=4, padx=5, pady=5, sticky="w")
        ToolTip(col_help, "The column index (starting at 0) that contains the website URL in your CSV file.")

        # --------------------
        # Row 2: wkhtmltopdf Path
        tk.Label(self.root, text="wkhtmltopdf Path:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        self.wk_entry = tk.Entry(self.root, width=50)
        self.wk_entry.grid(row=2, column=1, padx=5, pady=5)
        default_path = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe" if os.name == "nt" else "/usr/local/bin/wkhtmltopdf"
        self.wk_entry.insert(0, default_path)
        tk.Button(self.root, text="Browse", command=self.browse_wkhtmltopdf).grid(row=2, column=2, padx=5, pady=5)

        # --------------------
        # Row 3: Output Mode Selection
        self.output_mode = tk.StringVar(value="merged")
        mode_frame = tk.Frame(self.root)
        mode_frame.grid(row=3, column=0, columnspan=5, pady=5)
        tk.Label(mode_frame, text="Output Mode:").pack(side=tk.LEFT)
        tk.Radiobutton(mode_frame, text="Merge all into one PDF", variable=self.output_mode, value="merged",
                       command=self.toggle_output_options).pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(mode_frame, text="Save each webpage separately", variable=self.output_mode, value="separate",
                       command=self.toggle_output_options).pack(side=tk.LEFT, padx=5)

        # --------------------
        # Row 4: Output Options (Merged vs. Separate)
        # Merged mode widgets:
        self.output_pdf_label = tk.Label(self.root, text="Output PDF:")
        self.output_pdf_label.grid(row=4, column=0, sticky="e", padx=5, pady=5)
        self.output_pdf_entry = tk.Entry(self.root, width=50)
        self.output_pdf_entry.grid(row=4, column=1, padx=5, pady=5)
        self.output_pdf_entry.insert(0, "merged_output.pdf")
        self.output_pdf_browse_button = tk.Button(self.root, text="Browse", command=self.browse_save_pdf)
        self.output_pdf_browse_button.grid(row=4, column=2, padx=5, pady=5)

        # Separate mode widgets:
        self.output_dir_label = tk.Label(self.root, text="Output Directory:")
        self.output_dir_entry = tk.Entry(self.root, width=50)
        self.output_dir_browse_button = tk.Button(self.root, text="Browse", command=self.browse_output_dir)
        self.naming_label = tk.Label(self.root, text="Naming Scheme:")
        self.naming_scheme = tk.StringVar(value="Sequential with Timestamp")
        self.naming_combo = ttk.Combobox(self.root, textvariable=self.naming_scheme,
                                         values=["Sequential with Timestamp", "Website Domain"], state="readonly")
        self.naming_hint = tk.Label(self.root, text="(Website Domain: name PDFs after website domain; duplicates get index)")
        # Place separate mode widgets but hide them initially.
        self.output_dir_label.grid(row=4, column=0, sticky="e", padx=5, pady=5)
        self.output_dir_entry.grid(row=4, column=1, padx=5, pady=5)
        self.output_dir_browse_button.grid(row=4, column=2, padx=5, pady=5)
        self.naming_label.grid(row=5, column=0, sticky="e", padx=5, pady=5)
        self.naming_combo.grid(row=5, column=1, padx=5, pady=5, sticky="w")
        self.naming_hint.grid(row=5, column=2, padx=5, pady=5, sticky="w")
        self.hide_separate_mode_widgets()

        # --------------------
        # Row 6: Toggle Advanced Options Button
        self.toggle_adv_button = tk.Button(self.root, text="Hide Advanced Options", command=self.toggle_advanced)
        self.toggle_adv_button.grid(row=6, column=0, columnspan=5, pady=5)

        # --------------------
        # Row 7: Advanced Options (wkhtmltopdf settings & Concurrency)
        self.advanced_frame = tk.LabelFrame(self.root, text="Advanced Options")
        self.advanced_frame.grid(row=7, column=0, columnspan=5, padx=5, pady=5, sticky="ew")
        # wkhtmltopdf Options:
        tk.Label(self.advanced_frame, text="Page Size:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.page_size_var = tk.StringVar(value="A4")
        self.page_size_option = ttk.Combobox(self.advanced_frame, textvariable=self.page_size_var,
                                             values=["A4", "Letter", "Legal"], state="readonly")
        self.page_size_option.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        tk.Label(self.advanced_frame, text="Orientation:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.orientation_var = tk.StringVar(value="Portrait")
        self.orientation_option = ttk.Combobox(self.advanced_frame, textvariable=self.orientation_var,
                                               values=["Portrait", "Landscape"], state="readonly")
        self.orientation_option.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        # Margins:
        tk.Label(self.advanced_frame, text="Margin Top (mm):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.margin_top_entry = tk.Entry(self.advanced_frame, width=10)
        self.margin_top_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.margin_top_entry.insert(0, "10")
        tk.Label(self.advanced_frame, text="Margin Bottom (mm):").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.margin_bottom_entry = tk.Entry(self.advanced_frame, width=10)
        self.margin_bottom_entry.grid(row=1, column=3, padx=5, pady=5, sticky="w")
        self.margin_bottom_entry.insert(0, "10")
        tk.Label(self.advanced_frame, text="Margin Left (mm):").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.margin_left_entry = tk.Entry(self.advanced_frame, width=10)
        self.margin_left_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self.margin_left_entry.insert(0, "10")
        tk.Label(self.advanced_frame, text="Margin Right (mm):").grid(row=2, column=2, padx=5, pady=5, sticky="e")
        self.margin_right_entry = tk.Entry(self.advanced_frame, width=10)
        self.margin_right_entry.grid(row=2, column=3, padx=5, pady=5, sticky="w")
        self.margin_right_entry.insert(0, "10")
        # Concurrency: Max Workers
        tk.Label(self.advanced_frame, text="Max Workers:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.max_workers_spin = tk.Spinbox(self.advanced_frame, from_=1, to=16, width=5)
        self.max_workers_spin.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        self.max_workers_spin.delete(0, tk.END)
        self.max_workers_spin.insert(0, "3")
        tk.Label(self.advanced_frame, text="(More workers = faster but more CPU intensive)").grid(row=3, column=2, columnspan=2, padx=5, pady=5, sticky="w")

        # --------------------
        # Row 8: Progress Bar & ETA
        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=500, mode="determinate")
        self.progress_bar.grid(row=8, column=0, columnspan=5, padx=5, pady=5)
        self.eta_label = tk.Label(self.root, text="Estimated Time Remaining: N/A")
        self.eta_label.grid(row=9, column=0, columnspan=5, padx=5, pady=5)

        # --------------------
        # Row 10: Start & Cancel Buttons
        self.start_button = tk.Button(self.root, text="Start Conversion", command=self.start_conversion)
        self.start_button.grid(row=10, column=0, columnspan=3, pady=10)
        self.cancel_button = tk.Button(self.root, text="Cancel", command=self.cancel_conversion, state=tk.DISABLED)
        self.cancel_button.grid(row=10, column=3, columnspan=2, pady=10)

        # --------------------
        # Row 11: Log Output Area
        self.text_box = tk.Text(self.root, width=80, height=15, state=tk.DISABLED)
        self.text_box.grid(row=11, column=0, columnspan=5, padx=5, pady=5)

    # --------------------
    # Toggle Advanced Options Visibility
    def toggle_advanced(self):
        if self.advanced_visible.get():
            self.advanced_frame.grid_remove()
            self.toggle_adv_button.config(text="Show Advanced Options")
            self.advanced_visible.set(False)
        else:
            self.advanced_frame.grid()
            self.toggle_adv_button.config(text="Hide Advanced Options")
            self.advanced_visible.set(True)

    # --------------------
    # Utility Functions: File Dialogs
    def browse_csv(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if filename:
            self.csv_entry.delete(0, tk.END)
            self.csv_entry.insert(0, filename)

    def browse_wkhtmltopdf(self):
        filetypes = [("Executable", "*.exe")] if os.name == "nt" else None
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            self.wk_entry.delete(0, tk.END)
            self.wk_entry.insert(0, filename)

    def browse_save_pdf(self):
        filename = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")],
                                                initialfile="merged_output.pdf")
        if filename:
            self.output_pdf_entry.delete(0, tk.END)
            self.output_pdf_entry.insert(0, filename)

    def browse_output_dir(self):
        directory = filedialog.askdirectory()
        if directory:
            self.output_dir_entry.delete(0, tk.END)
            self.output_dir_entry.insert(0, directory)

    def hide_separate_mode_widgets(self):
        self.output_dir_label.grid_remove()
        self.output_dir_entry.grid_remove()
        self.output_dir_browse_button.grid_remove()
        self.naming_label.grid_remove()
        self.naming_combo.grid_remove()
        self.naming_hint.grid_remove()

    def show_separate_mode_widgets(self):
        self.output_dir_label.grid()
        self.output_dir_entry.grid()
        self.output_dir_browse_button.grid()
        self.naming_label.grid()
        self.naming_combo.grid()
        self.naming_hint.grid()

    def toggle_output_options(self):
        if self.output_mode.get() == "merged":
            self.output_pdf_label.grid()
            self.output_pdf_entry.grid()
            self.output_pdf_browse_button.grid()
            self.hide_separate_mode_widgets()
        else:
            self.output_pdf_label.grid_remove()
            self.output_pdf_entry.grid_remove()
            self.output_pdf_browse_button.grid_remove()
            self.show_separate_mode_widgets()

    # --------------------
    # Logging Helpers
    def log(self, message):
        self.text_box.config(state=tk.NORMAL)
        self.text_box.insert(tk.END, message + "\n")
        self.text_box.config(state=tk.DISABLED)
        self.text_box.see(tk.END)
        logging.info(message)

    def log_exception(self, message):
        self.log(message)
        logging.exception(message)

    # --------------------
    # ETA Calculation
    def update_eta(self, elapsed, completed, total):
        if completed == 0:
            eta = "Calculating..."
        else:
            avg_time = elapsed / completed
            remaining = total - completed
            eta_seconds = remaining * avg_time
            eta = f"{int(eta_seconds)} sec"
        self.eta_label.config(text=f"Estimated Time Remaining: {eta}")

    # --------------------
    # Worker function to process a single URL.
    def process_url(self, url, index, config, options, output_path):
        if cancel_event.is_set():
            return None
        try:
            pdfkit.from_url(url, output_path, configuration=config, options=options)
            return (index, output_path, None)
        except Exception as e:
            logging.exception(f"Error processing URL {url}")
            return (index, output_path, str(e))

    # --------------------
    # Start Conversion (runs in a separate thread)
    def start_conversion(self):
        cancel_event.clear()
        self.start_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)
        self.text_box.config(state=tk.NORMAL)
        self.text_box.delete("1.0", tk.END)
        self.text_box.config(state=tk.DISABLED)
        self.progress_bar["value"] = 0
        self.eta_label.config(text="Estimated Time Remaining: Calculating...")
        self.start_time = time.time()

        # Validate CSV file.
        csv_file = self.csv_entry.get().strip()
        if not os.path.isfile(csv_file):
            messagebox.showerror("Error", "Please select a valid CSV file.")
            self.start_button.config(state=tk.NORMAL)
            self.cancel_button.config(state=tk.DISABLED)
            return

        # Validate wkhtmltopdf executable.
        wk_path = self.wk_entry.get().strip()
        if not os.path.isfile(wk_path):
            messagebox.showerror("Error", "Please select a valid wkhtmltopdf executable path.")
            self.start_button.config(state=tk.NORMAL)
            self.cancel_button.config(state=tk.DISABLED)
            return

        try:
            config_obj = pdfkit.configuration(wkhtmltopdf=wk_path)
        except Exception as e:
            messagebox.showerror("Error", f"Error configuring pdfkit: {e}")
            self.start_button.config(state=tk.NORMAL)
            self.cancel_button.config(state=tk.DISABLED)
            return

        # Read CSV and extract URLs.
        urls = []
        try:
            with open(csv_file, newline="", encoding="utf-8") as f:
                reader = csv.reader(f)
                rows = list(reader)
                if self.csv_header_var.get() and rows:
                    rows = rows[1:]
                col_index = int(self.csv_column_entry.get().strip())
                for row in rows:
                    if len(row) > col_index:
                        urls.append(row[col_index].strip())
        except Exception as e:
            messagebox.showerror("Error", f"Error reading CSV file: {e}")
            self.start_button.config(state=tk.NORMAL)
            self.cancel_button.config(state=tk.DISABLED)
            return

        if not urls:
            messagebox.showinfo("Info", "No URLs found in CSV file.")
            self.start_button.config(state=tk.NORMAL)
            self.cancel_button.config(state=tk.DISABLED)
            return

        self.log(f"Found {len(urls)} URLs.")
        total_urls = len(urls)
        self.progress_bar["maximum"] = total_urls

        # Prepare wkhtmltopdf options.
        options = {
            "page-size": self.page_size_var.get(),
            "orientation": self.orientation_var.get(),
            "margin-top": self.margin_top_entry.get(),
            "margin-bottom": self.margin_bottom_entry.get(),
            "margin-left": self.margin_left_entry.get(),
            "margin-right": self.margin_right_entry.get()
        }
        # Validate numeric fields (margins).
        try:
            float(options["margin-top"])
            float(options["margin-bottom"])
            float(options["margin-left"])
            float(options["margin-right"])
        except ValueError:
            messagebox.showerror("Error", "Margins must be numeric values.")
            self.start_button.config(state=tk.NORMAL)
            self.cancel_button.config(state=tk.DISABLED)
            return

        mode = self.output_mode.get()
        if mode == "merged":
            temp_folder = os.path.join(os.getcwd(), "temp_pdfs")
            os.makedirs(temp_folder, exist_ok=True)
            output_paths = {}
        else:
            output_dir = self.output_dir_entry.get().strip()
            if not os.path.isdir(output_dir):
                self.log("Output directory invalid. Using current directory.")
                output_dir = os.getcwd()
            output_paths = {}

        # Get max workers from spinbox.
        try:
            max_workers = int(self.max_workers_spin.get())
        except ValueError:
            messagebox.showerror("Error", "Max Workers must be a numeric value.")
            self.start_button.config(state=tk.NORMAL)
            self.cancel_button.config(state=tk.DISABLED)
            return

        # Start processing URLs concurrently.
        futures = []
        executor = concurrent.futures.ThreadPoolExecutor(max_workers=max_workers)
        for i, url in enumerate(urls, start=1):
            if mode == "merged":
                output_path = os.path.join(temp_folder, f"page_{i}.pdf")
            else:
                # Choose naming scheme.
                if self.naming_scheme.get() == "Website Domain":
                    try:
                        parsed = urlparse(url)
                        domain = parsed.netloc.replace("www.", "")
                        base_name = domain if domain else "page"
                    except Exception:
                        base_name = "page"
                    output_path = os.path.join(output_dir, f"{base_name}_{i}.pdf")
                else:
                    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                    output_path = os.path.join(output_dir, f"page_{i}_{timestamp}.pdf")
            output_paths[i] = output_path
            future = executor.submit(self.process_url, url, i, config_obj, options, output_path)
            futures.append(future)

        def check_futures():
            completed = sum(1 for f in futures if f.done())
            elapsed = time.time() - self.start_time
            self.update_eta(elapsed, completed, total_urls)
            self.progress_bar["value"] = completed
            if completed < total_urls and not cancel_event.is_set():
                self.root.after(500, check_futures)
            else:
                executor.shutdown(wait=False)
                self.finish_conversion(mode, output_paths)
                self.start_button.config(state=tk.NORMAL)
                self.cancel_button.config(state=tk.DISABLED)

        self.root.after(500, check_futures)

    # --------------------
    # Finalizing Conversion & Merging PDFs
    def finish_conversion(self, mode, output_paths):
        results = []
        for index, path in output_paths.items():
            results.append((index, path))
        results.sort(key=lambda x: x[0])
        if mode == "merged":
            merger = PdfMerger()
            for idx, pdf_path in results:
                if os.path.exists(pdf_path):
                    merger.append(pdf_path)
            output_pdf = self.output_pdf_entry.get().strip()
            try:
                merger.write(output_pdf)
                merger.close()
                self.log(f"Merged PDF created at {output_pdf}")
                messagebox.showinfo("Success", f"Merged PDF created at {output_pdf}")
            except Exception as e:
                self.log(f"Error merging PDFs: {e}")
                logging.exception("Merging error")
                messagebox.showerror("Error", f"Error merging PDFs: {e}")
            # Cleanup temporary PDFs.
            for idx, pdf_path in results:
                try:
                    os.remove(pdf_path)
                except Exception:
                    pass
            try:
                os.rmdir(os.path.dirname(list(output_paths.values())[0]))
            except Exception:
                pass
        elif mode == "separate":
            self.log("Individual PDFs created successfully.")
            messagebox.showinfo("Success", "Individual PDFs have been created.")

    # --------------------
    # Cancel Button Handler
    def cancel_conversion(self):
        cancel_event.set()
        self.log("Conversion cancelled by user.")
        self.start_button.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)

# =============================================================================
# Main Entry Point
# =============================================================================
def main():
    root = tk.Tk()
    app = WebPage2PDFBundle(root)
    root.mainloop()

if __name__ == "__main__":
    main()
