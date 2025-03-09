import os
import sys
import csv
import time
import threading
import subprocess
import logging
import concurrent.futures

# Try to import required modules; if missing, they will be installed during bootstrap.
try:
    import pdfkit
    from PyPDF2 import PdfMerger
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError:
    pass  # Bootstrap will handle installation

# =============================================================================
# Bootstrap Virtual Environment & Dependency Installation
# =============================================================================
def bootstrap_venv():
    """
    Checks if running inside a virtual environment. If not, creates one in a folder 'venv',
    installs dependencies, and relaunches the script from the virtual environment.
    """
    if sys.prefix != sys.base_prefix:
        return  # Already inside a virtual environment.

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

    # Relaunch the script using the virtual environment's python.
    print("Relaunching script in virtual environment...")
    subprocess.check_call([python_executable, __file__])
    sys.exit()

# Only bootstrap if not already in a venv.
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
# Main Application Class
# =============================================================================
class WebPage2PDFBundle:
    def __init__(self, root):
        self.root = root
        self.root.title("WebPage2PDF Bundle")
        self.start_time = None
        self.setup_gui()

    def setup_gui(self):
        # --------------------
        # Row 0: CSV File Selection
        tk.Label(self.root, text="CSV File:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.csv_entry = tk.Entry(self.root, width=50)
        self.csv_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(self.root, text="Browse", command=self.browse_csv).grid(row=0, column=2, padx=5, pady=5)

        # --------------------
        # Row 1: wkhtmltopdf Path
        tk.Label(self.root, text="wkhtmltopdf Path:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.wk_entry = tk.Entry(self.root, width=50)
        self.wk_entry.grid(row=1, column=1, padx=5, pady=5)
        default_path = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe" if os.name == "nt" else "/usr/local/bin/wkhtmltopdf"
        self.wk_entry.insert(0, default_path)
        tk.Button(self.root, text="Browse", command=self.browse_wkhtmltopdf).grid(row=1, column=2, padx=5, pady=5)

        # --------------------
        # Row 2: Output Mode Selection
        self.output_mode = tk.StringVar(value="merged")
        mode_frame = tk.Frame(self.root)
        mode_frame.grid(row=2, column=0, columnspan=3, pady=5)
        tk.Label(mode_frame, text="Output Mode: ").pack(side=tk.LEFT)
        tk.Radiobutton(mode_frame, text="Merge all into one PDF", variable=self.output_mode, value="merged",
                       command=self.toggle_output_options).pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(mode_frame, text="Save each webpage separately", variable=self.output_mode, value="separate",
                       command=self.toggle_output_options).pack(side=tk.LEFT, padx=5)

        # --------------------
        # Row 3: Output Options
        self.output_pdf_label = tk.Label(self.root, text="Output PDF:")
        self.output_pdf_label.grid(row=3, column=0, sticky="e", padx=5, pady=5)
        self.output_pdf_entry = tk.Entry(self.root, width=50)
        self.output_pdf_entry.grid(row=3, column=1, padx=5, pady=5)
        self.output_pdf_entry.insert(0, "merged_output.pdf")
        self.output_pdf_browse_button = tk.Button(self.root, text="Browse", command=self.browse_save_pdf)
        self.output_pdf_browse_button.grid(row=3, column=2, padx=5, pady=5)

        self.output_dir_label = tk.Label(self.root, text="Output Directory:")
        self.output_dir_entry = tk.Entry(self.root, width=50)
        self.output_dir_browse_button = tk.Button(self.root, text="Browse", command=self.browse_output_dir)
        # Initially hide separate mode fields.
        self.output_dir_label.grid(row=3, column=0, sticky="e", padx=5, pady=5)
        self.output_dir_entry.grid(row=3, column=1, padx=5, pady=5)
        self.output_dir_browse_button.grid(row=3, column=2, padx=5, pady=5)
        self.output_dir_label.grid_remove()
        self.output_dir_entry.grid_remove()
        self.output_dir_browse_button.grid_remove()

        # --------------------
        # Row 4: Advanced Options (Collapsible)
        self.advanced_frame = tk.LabelFrame(self.root, text="Advanced Options (wkhtmltopdf settings)")
        self.advanced_frame.grid(row=4, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        # Page Size
        tk.Label(self.advanced_frame, text="Page Size:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.page_size_var = tk.StringVar(value="A4")
        self.page_size_option = ttk.Combobox(self.advanced_frame, textvariable=self.page_size_var,
                                             values=["A4", "Letter", "Legal"])
        self.page_size_option.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        # Orientation
        tk.Label(self.advanced_frame, text="Orientation:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.orientation_var = tk.StringVar(value="Portrait")
        self.orientation_option = ttk.Combobox(self.advanced_frame, textvariable=self.orientation_var,
                                               values=["Portrait", "Landscape"])
        self.orientation_option.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        # Margins
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

        # --------------------
        # Row 5: Progress Bar & ETA
        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=500, mode="determinate")
        self.progress_bar.grid(row=5, column=0, columnspan=3, padx=5, pady=5)
        self.eta_label = tk.Label(self.root, text="Estimated Time Remaining: N/A")
        self.eta_label.grid(row=6, column=0, columnspan=3, padx=5, pady=5)

        # --------------------
        # Row 7: Start & Cancel Buttons
        self.start_button = tk.Button(self.root, text="Start Conversion", command=self.start_conversion)
        self.start_button.grid(row=7, column=0, columnspan=2, pady=10)
        self.cancel_button = tk.Button(self.root, text="Cancel", command=self.cancel_conversion, state=tk.DISABLED)
        self.cancel_button.grid(row=7, column=2, pady=10)

        # --------------------
        # Row 8: Log Output Area
        self.text_box = tk.Text(self.root, width=80, height=15, state=tk.DISABLED)
        self.text_box.grid(row=8, column=0, columnspan=3, padx=5, pady=5)

    # --------------------
    # File Dialogs
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

    def toggle_output_options(self):
        if self.output_mode.get() == "merged":
            self.output_pdf_label.grid()
            self.output_pdf_entry.grid()
            self.output_pdf_browse_button.grid()
            self.output_dir_label.grid_remove()
            self.output_dir_entry.grid_remove()
            self.output_dir_browse_button.grid_remove()
        else:
            self.output_pdf_label.grid_remove()
            self.output_pdf_entry.grid_remove()
            self.output_pdf_browse_button.grid_remove()
            self.output_dir_label.grid()
            self.output_dir_entry.grid()
            self.output_dir_browse_button.grid()

    # --------------------
    # Logging Helper
    def log(self, message):
        self.text_box.config(state=tk.NORMAL)
        self.text_box.insert(tk.END, message + "\n")
        self.text_box.config(state=tk.DISABLED)
        self.text_box.see(tk.END)
        logging.info(message)

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
    # Conversion Worker
    def process_url(self, url, index, config, options, output_path):
        if cancel_event.is_set():
            return None
        try:
            pdfkit.from_url(url, output_path, configuration=config, options=options)
            return (index, output_path, None)
        except Exception as e:
            return (index, output_path, str(e))

    # --------------------
    # Start Conversion (Runs in Separate Thread)
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

        csv_file = self.csv_entry.get().strip()
        wk_path = self.wk_entry.get().strip()
        mode = self.output_mode.get()

        if not os.path.isfile(csv_file):
            messagebox.showerror("Error", "Please select a valid CSV file.")
            self.start_button.config(state=tk.NORMAL)
            self.cancel_button.config(state=tk.DISABLED)
            return

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

        # Read URLs from CSV.
        urls = []
        try:
            with open(csv_file, newline="", encoding="utf-8") as f:
                reader = csv.reader(f)
                for row in reader:
                    if row:
                        urls.append(row[0].strip())
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

        # Prepare wkhtmltopdf advanced options.
        options = {
            "page-size": self.page_size_var.get(),
            "orientation": self.orientation_var.get(),
            "margin-top": self.margin_top_entry.get(),
            "margin-bottom": self.margin_bottom_entry.get(),
            "margin-left": self.margin_left_entry.get(),
            "margin-right": self.margin_right_entry.get()
        }

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

        max_workers = min(4, total_urls) if total_urls > 0 else 1
        futures = []
        executor = concurrent.futures.ThreadPoolExecutor(max_workers=max_workers)
        for i, url in enumerate(urls, start=1):
            if mode == "merged":
                output_path = os.path.join(temp_folder, f"page_{i}.pdf")
            else:
                output_path = os.path.join(output_dir, f"page_{i}.pdf")
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
    # Finish Conversion & Merging
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
                messagebox.showerror("Error", f"Error merging PDFs: {e}")
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
