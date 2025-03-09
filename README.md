# WebPage2PDF Bundle


**WebPage2PDF Bundle** is a Windows‑focused, user‑friendly tool that converts website URLs from a CSV file into PDF documents. Users can choose to merge all webpages into one PDF or save each webpage as its own PDF. Advanced wkhtmltopdf settings (page size, orientation, margins) are available without any code changes. The tool includes a responsive GUI with a progress bar, estimated time remaining, detailed logging, and cancel functionality.

## Features

- **Automatic Setup:**  
  – When first run, the tool automatically creates a virtual environment and installs all required dependencies.  
  – No manual dependency installation is needed.

- **Intuitive GUI:**  
  – Select your CSV file and wkhtmltopdf executable using file dialogs.  
  – Choose between merging all webpages into one PDF or saving separate PDFs.
  – Advanced options to configure wkhtmltopdf parameters (page size, orientation, margins).

- **Concurrent Processing:**  
  – Converts webpages concurrently while updating a progress bar and estimating remaining time.
  
- **Robust Logging & Cancel:**  
  – Detailed log output is displayed on‑screen and saved to `conversion.log`.
  – Easily cancel the conversion process at any time.

## Requirements

- **Windows OS**  
- **Python 3.6+** (the tool will set up its own virtual environment)

## How to Use

1.  **Install wkhtmltopdf:** Download and install wkhtmltopdf from wkhtmltopdf.org/downloads.html.The tool’s default path assumes:

2. **Download the Files:**  
   – `webpage2pdf_bundle.py`  
   – `README.md`

3. **Run the Tool:**  
   Simply double‑click `webpage2pdf_bundle.py` or run it from the command prompt:

   ```bash
   python webpage2pdf_bundle.py
    On first run, the script will create a virtual environment (in the venv folder), install required dependencies, and relaunch itself automatically.

4. **Follow the GUI Prompts:**

    *   **CSV File:** Browse and select your CSV file containing one website URL per row.
        
    *   **wkhtmltopdf Path:** Verify or change the path to the wkhtmltopdf executable (default is set for Windows).

    *   **Output Mode:** Choose whether to merge all webpages into one PDF or save individual PDFs.
        
    *   **Advanced Options:** Adjust page size, orientation, and margins as needed.

        
    *   **Start Conversion:** Click the "Start Conversion" button and monitor progress. Use the "Cancel" button if needed.
        
4.  **Output:**
    
    *   For merged mode, the final PDF is saved as specified.
        

    *   For separate mode, individual PDFs are saved in the chosen output directory.


     

