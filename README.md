### How to Set Up and Use WebPage2PDF Bundle

1.  pip install pdfkit PyPDF2   (Linux users may need to install tkinter—often via your package manager, e.g., sudo apt-get install python3-tk.)
    
2.  **Install wkhtmltopdf:**Download and install wkhtmltopdf from wkhtmltopdf.org/downloads.html.The tool’s default path assumes:
    
    *   **Windows:** C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe
        
    *   **macOS/Linux:** /usr/local/bin/wkhtmltopdfYou can change this path with the provided “Browse” button if needed.
        
3.  **Prepare Your CSV File:**Create a CSV file containing one website URL per row(do it yourself or you can search and use a tool online).
    
4.  then run in your terminal with  python main.py
    
5.  **Follow the GUI Prompts:**
    
    *   **CSV File:** Click “Browse” to select your CSV file.
        
    *   **wkhtmltopdf Path:** Verify or change the path to the wkhtmltopdf executable.
        
    *   **Output Mode:** Choose “Merge all into one PDF” (default) or “Save each webpage separately.”
        
        *   For merged mode, specify the desired output PDF file name (with “Browse” to choose a save-as location).
            
        *   For separate mode, choose the output directory where individual PDFs will be saved.
            
    *   Click “Start Conversion.” Watch the progress and log for updates.
