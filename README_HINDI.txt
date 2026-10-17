TENDER TECHNICAL BID COMPARATIVE STATEMENT SOFTWARE
Version: Starter 1.0

WHAT IT DOES
- Upload multiple agency-wise ZIP files.
- Reads text from PDF and DOCX documents inside ZIP files.
- Attempts to identify DD, EMD/FDR, PAN, GST, electrical contractor licence, solvency and agency registration details.
- Lets the user correct/edit every extracted field.
- Creates a landscape Word comparative statement with signature spaces.
- Does not include a "3A Valid" column.

IMPORTANT LIMITATIONS
- This is an initial working version, not a certified tender scrutiny system.
- OCR is optional and scanned documents may not be read unless Tesseract OCR and Poppler are installed.
- The extraction rules are heuristic. Verify all values, dates, amounts and legal eligibility against original documents.
- Do not use automatic extraction alone to decide agency qualification.
- Best results: one agency per ZIP, PDFs with selectable text, clear file names.

INSTALL (WINDOWS)
1. Install Python 3.10+ from https://www.python.org/downloads/
   During installation, tick "Add Python to PATH".
2. Extract this ZIP into a folder.
3. Double-click install.bat and wait for completion.
4. Double-click run.bat.
5. Browser will open at http://localhost:8501.
6. Upload the agency technical-bid ZIP files.

OPTIONAL OCR
For scanned PDFs, install Tesseract OCR and Poppler separately and ensure their executables are on PATH. Install OCR packages by running:
  pip install pytesseract pdf2image
OCR language is currently English (eng). Gujarati OCR needs suitable Gujarati Tesseract language data and configuration.

FILES
- app.py: application
- requirements.txt: required Python packages
- install.bat: installs dependencies
- run.bat: launches the app

DATA / PRIVACY
The app is designed to run locally on your own computer. Do not expose the local app publicly or upload confidential tender documents to a third-party host without approval.
