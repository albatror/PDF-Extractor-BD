# PDF Travel Expense Extractor

A portable application for extracting travel expense data from PDF forms using OCR and text extraction.

## Features

- Extracts data from both text-based and scanned PDFs
- Supports A4 (4 pages) and A3 (2 pages) form formats
- Interactive PDF preview with annotation capabilities
- Automatic form type detection
- Excel export functionality
- Robust error handling for corrupted or unreadable PDFs

## Requirements

- Python 3.6+
- PyQt5
- PyPDF2
- pdfplumber
- pytesseract
- OpenCV
- pandas
- openpyxl

## Installation

1. Install Python (if not already installed)
2. Install required packages:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```
   python main.py
   ```

2. Select PDF files to process
3. Preview PDFs and annotate important areas
4. Extract data from forms
5. Export results to Excel

## Portability

This application is designed to work without administrator privileges:
- All dependencies are installed in user directory
- No system-wide installations required
- Portable Python and Tesseract OCR included (see installation notes)

## License

MIT License
