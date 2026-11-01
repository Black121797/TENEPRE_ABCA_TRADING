# Abaca Fiber Billing & Proforma Management System 🚀

**Engine Core Designed by: RAVEN TENEPRE**

## Project Overview

A Windows desktop utility for abaca receiving-report review, fiber-weight auditing, and proforma packet preparation. The application extracts available receiving-report text and uses local OCR to help identify report metadata and fiber weights. It calculates line amounts and the grand total as values are edited, then prepares a two-page PDF packet using the output-form template and the supplier sales invoice image.

The optional PDF-XChange workflow opens a temporary copy of the receiving report with page-one OCR options (English, high accuracy, searchable image). Complete OCR and save the result in PDF-XChange, then import the saved PDF in the application. This step is user-guided; the program does not automatically save PDF-XChange's OCR output.

## Fiber Price Configuration

Unit prices are loaded from `fiber_prices.txt`, located alongside the Python script. The management team can change a price using Notepad without editing Python source. Keep one `GRADE=PRICE` entry on each line:

```text
S2=95.00
I=89.00
G=72.00
AD1=92.00
S3=72.00
```

Prices are expressed in PHP per kilogram. The application has defaults for all five grades and reads overrides from this file.

## System Prerequisites (Windows Only)

- Windows 10 or later.
- Python 3.10 or later for running from source.
- PDF-XChange Editor for document preview and the optional guided OCR workflow. The default executable path used by the application is:

  `C:\Program Files\Tracker Software\PDF Editor\PDFXEdit.exe`

  The application also attempts to discover supported Editor executable names through the Windows Registry and allows manual location if automatic detection fails.
- Install PDF-XChange Editor from the [official PDF-XChange download page](https://pdf-xchange.com).
- The `OUTPUT FILE FORMAT.pdf` template must be present beside the script for packet export. PDF files are excluded by this repository's `.gitignore`; supply the approved template separately after cloning.

## Developer Workspace Installation Track

Run these commands from the project directory in Command Prompt:

```bat
# 1. Create and activate the isolated virtual environment
py -m venv .venv
.venv\Scripts\activate

# 2. Install dependencies declared in the requirements manifest
pip install -r requirements.txt

# 3. Launch the application
python fiber_report_packet.py
```

## Typical Workflow

1. Select the receiving-report PDF. The application opens it in PDF-XChange for preview and reads available report data.
2. Select the supplier Sales Invoice image (JPEG or PNG).
3. Review the extracted P.O. number, R.R. number, billing date, and grade weights. Correct or enter missing values in the highlighted fields.
4. If needed, use **Run PDF-XChange OCR (page 1)**, finish and save OCR in PDF-XChange, then use **Import saved OCR PDF**.
5. Confirm the configured prices and calculated totals, then choose **FINALIZE & EXPORT REPORT PACKET**.

The exported packet uses the output-format PDF as its first page and appends the supplier invoice image as its second page.

## Standalone Compilation & Deployment Packaging Guide

Install PyInstaller into the active virtual environment and build the executable:

```bat
pip install pyinstaller
pyinstaller --onefile --noconsole --name="Abaca_Fiber_Billing" fiber_report_packet.py
```

The executable is written to `dist\Abaca_Fiber_Billing.exe`. This packages the Python application and its Python dependencies; it does **not** make deployment dependency-free because PDF-XChange Editor must still be installed.

The command above does not bundle the external `fiber_prices.txt` configuration or `OUTPUT FILE FORMAT.pdf` template. Also, the current source resolves these files relative to `__file__`; a one-file PyInstaller build extracts bundled code to a temporary directory. Before production deployment, update resource-path handling for frozen builds and either bundle the required resources with PyInstaller's `--add-data` option or explicitly resolve them from a documented external folder. The output template is ignored by Git and must be obtained separately. Test the final package on a clean Windows machine.
