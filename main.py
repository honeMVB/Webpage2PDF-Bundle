import os
import csv
import threading
import pdfkit
from PyPDF2 import PdfMerger
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

# ----------------------
# Helper functions for thread-safe GUI updates
def safe_log(message):
    root.after(0, lambda: log(message))

def safe_update_progress(value):
    root.after(0, lambda: progress_bar.config(value=value))

# ----------------------
# Log function for the text widget
def log(message):
    text_box.config(state=tk.NORMAL)
    text_box.insert(tk.END, message + "\n")
    text_box.config(state=tk.DISABLED)
    text_box.see(tk.END)

# ----------------------
# File dialogs
def browse_csv():
    filename = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if filename:
        csv_entry.delete(0, tk.END)
        csv_entry.insert(0, filename)

def browse_wkhtmltopdf():
    filetypes = [("Executable", "*.exe")] if os.name == 'nt' else None
    filename = filedialog.askopenfilename(filetypes=filetypes)
    if filename:
        wk_entry.delete(0, tk.END)
        wk_entry.insert(0, filename)

def browse_output_dir():
    directory = filedialog.askdirectory()
    if directory:
        output_dir_entry.delete(0, tk.END)
        output_dir_entry.insert(0, directory)

# ----------------------
# Toggle output options based on mode
def toggle_output_options():
    if output_mode.get() == "merged":
        # Show merged output field; hide output directory field.
        output_pdf_label.grid()
        output_pdf_entry.grid()
        output_pdf_browse_button.grid()
        output_dir_label.grid_remove()
        output_dir_entry.grid_remove()
        output_dir_browse_button.grid_remove()
    else:
        # Show output directory field; hide merged output field.
        output_pdf_label.grid_remove()
        output_pdf_entry.grid_remove()
        output_pdf_browse_button.grid_remove()
        output_dir_label.grid()
        output_dir_entry.grid()
        output_dir_browse_button.grid()

# ----------------------
# The main conversion function (runs in a separate thread)
def process_conversion():
    csv_file = csv_entry.get().strip()
    wk_path = wk_entry.get().strip()
    mode = output_mode.get()

    # Validate CSV file
    if not os.path.isfile(csv_file):
        safe_log("Error: CSV file not found.")
        messagebox.showerror("Error", "Please select a valid CSV file.")
        return

    # Validate wkhtmltopdf executable
    if not os.path.isfile(wk_path):
        safe_log("Error: wkhtmltopdf executable not found.")
        messagebox.showerror("Error", "Please select a valid wkhtmltopdf executable path.")
        return

    try:
        config = pdfkit.configuration(wkhtmltopdf=wk_path)
    except Exception as e:
        safe_log(f"Error configuring pdfkit: {e}")
        messagebox.showerror("Error", f"Error configuring pdfkit: {e}")
        return

    # Read URLs from the CSV file.
    urls = []
    try:
        with open(csv_file, newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            for row in reader:
                if row:
                    urls.append(row[0].strip())
    except Exception as e:
        safe_log(f"Error reading CSV file: {e}")
        messagebox.showerror("Error", f"Error reading CSV file: {e}")
        return

    if not urls:
        safe_log("No URLs found in the CSV file.")
        messagebox.showinfo("Info", "No URLs found in the CSV file.")
        return

    safe_log(f"Found {len(urls)} URLs.")
    progress_bar["maximum"] = len(urls)
    progress_bar["value"] = 0

    if mode == "merged":
        # Merged mode: create a temporary folder and store individual PDFs for merging.
        temp_folder = "temp_pdfs"
        os.makedirs(temp_folder, exist_ok=True)
        pdf_files = []
    else:
        # Separate mode: output directory is used.
        output_dir = output_dir_entry.get().strip()
        if not os.path.isdir(output_dir):
            safe_log("Output directory is invalid. Using current directory.")
            output_dir = os.getcwd()

    # Process each URL.
    for i, url in enumerate(urls, start=1):
        safe_log(f"Processing {url} ...")
        if mode == "merged":
            output_path = os.path.join(temp_folder, f"page_{i}.pdf")
        else:
            # Save individual PDFs with a sequential naming.
            output_path = os.path.join(output_dir, f"page_{i}.pdf")
        try:
            pdfkit.from_url(url, output_path, configuration=config)
            safe_log(f"Saved PDF for {url}")
            if mode == "merged":
                pdf_files.append(output_path)
        except Exception as e:
            safe_log(f"Failed to convert {url}: {e}")
        safe_update_progress(i)

    # Finalizing: either merge PDFs or finish.
    if mode == "merged" and pdf_files:
        merger = PdfMerger()
        for pdf in pdf_files:
            merger.append(pdf)
        output_pdf = output_pdf_entry.get().strip()
        try:
            merger.write(output_pdf)
            merger.close()
            safe_log(f"Merged PDF created at {output_pdf}")
            messagebox.showinfo("Success", f"Merged PDF created at {output_pdf}")
        except Exception as e:
            safe_log(f"Error merging PDFs: {e}")
            messagebox.showerror("Error", f"Error merging PDFs: {e}")
        # Cleanup temporary PDFs.
        for pdf in pdf_files:
            try:
                os.remove(pdf)
            except Exception:
                pass
        try:
            os.rmdir(temp_folder)
        except Exception:
            pass
    elif mode == "separate":
        safe_log("Individual PDFs created successfully.")
        messagebox.showinfo("Success", "Individual PDFs have been created.")
    else:
        safe_log("No PDFs were created.")

# ----------------------
# Start conversion on a separate thread
def start_conversion():
    t = threading.Thread(target=process_conversion)
    t.start()

# ----------------------
# Set up the GUI
root = tk.Tk()
root.title("WebPage2PDF Bundle")

# --- Row 0: CSV File Selection ---
tk.Label(root, text="CSV File:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
csv_entry = tk.Entry(root, width=50)
csv_entry.grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="Browse", command=browse_csv).grid(row=0, column=2, padx=5, pady=5)

# --- Row 1: wkhtmltopdf Path ---
tk.Label(root, text="wkhtmltopdf Path:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
wk_entry = tk.Entry(root, width=50)
wk_entry.grid(row=1, column=1, padx=5, pady=5)
default_path = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe" if os.name == 'nt' else "/usr/local/bin/wkhtmltopdf"
wk_entry.insert(0, default_path)
tk.Button(root, text="Browse", command=browse_wkhtmltopdf).grid(row=1, column=2, padx=5, pady=5)

# --- Row 2: Output Mode Selection ---
output_mode = tk.StringVar(value="merged")
mode_frame = tk.Frame(root)
mode_frame.grid(row=2, column=0, columnspan=3, pady=5)
tk.Label(mode_frame, text="Output Mode: ").pack(side=tk.LEFT)
tk.Radiobutton(mode_frame, text="Merge all into one PDF", variable=output_mode, value="merged",
               command=toggle_output_options).pack(side=tk.LEFT, padx=5)
tk.Radiobutton(mode_frame, text="Save each webpage separately", variable=output_mode, value="separate",
               command=toggle_output_options).pack(side=tk.LEFT, padx=5)

# --- Row 3: Output Options (Merged or Separate) ---
# For merged mode:
output_pdf_label = tk.Label(root, text="Output PDF:")
output_pdf_label.grid(row=3, column=0, sticky="e", padx=5, pady=5)
output_pdf_entry = tk.Entry(root, width=50)
output_pdf_entry.grid(row=3, column=1, padx=5, pady=5)
output_pdf_entry.insert(0, "merged_output.pdf")
output_pdf_browse_button = tk.Button(root, text="Browse", command=lambda: filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")], initialfile="merged_output.pdf"))
output_pdf_browse_button.grid(row=3, column=2, padx=5, pady=5)

# For separate mode (initially hidden)
output_dir_label = tk.Label(root, text="Output Directory:")
output_dir_entry = tk.Entry(root, width=50)
output_dir_browse_button = tk.Button(root, text="Browse", command=browse_output_dir)
# Place in grid but hide for now.
output_dir_label.grid(row=3, column=0, sticky="e", padx=5, pady=5)
output_dir_entry.grid(row=3, column=1, padx=5, pady=5)
output_dir_browse_button.grid(row=3, column=2, padx=5, pady=5)
output_dir_label.grid_remove()
output_dir_entry.grid_remove()
output_dir_browse_button.grid_remove()

# --- Row 4: Progress Bar ---
progress_bar = ttk.Progressbar(root, orient="horizontal", length=500, mode="determinate")
progress_bar.grid(row=4, column=0, columnspan=3, padx=5, pady=5)

# --- Row 5: Start Conversion Button ---
tk.Button(root, text="Start Conversion", command=start_conversion).grid(row=5, column=0, columnspan=3, pady=10)

# --- Row 6: Log Output Area ---
text_box = tk.Text(root, width=80, height=15, state=tk.DISABLED)
text_box.grid(row=6, column=0, columnspan=3, padx=5, pady=5)

root.mainloop()
