
# PDF Data Extractor and Excel Writer

This repository contains a Python-based solution designed to extract specific data from multiple PDF files and compile the extracted information into an Excel file. Utilizing the PyMuPDF library for PDF processing and the openpyxl library for Excel file handling, this project automates the extraction and organization of data from a folder of PDF files into a structured Excel spreadsheet.

## Features

- PDF Data Extraction: Extracts specified fields from PDFs using regex patterns.
- Data Transformation: Splits a location field into district, taluka, and village fields.
- Excel Writing: Writes the extracted data to an Excel file with appropriate column headings.
- Indexing: Adds an index to each row in the Excel file for easy reference.


## Prerequisites

- Python 3.6 or higher
- The following Python libraries:
    - PyMuPDF (fitz)
    - openpyxl


## Installation

## Installation

Clone the Repository:

```bash
  git clone https://github.com/rushigiri11/PDF_to_Excel_Data_Extraction.git
  cd PDF_to_Excel_Data_Extraction
```


Install the Required Libraries:

```bash
  pip install pymupdf openpyxl
  ```

Project Structure
```bash
pdf-data-extractor/
├── pdfdata/
│   ├── file1.pdf
│   ├── file2.pdf
│   └── file3.pdf
├── extract_pdf_data.py
└── README.md
  ```

  - pdfdata/: Directory containing PDF files to be processed.
  - extract_pdf_data.py: Main script for extracting data from PDFs and writing to an Excel file.
  - README.md: Project description and instructions.







## Usage
 Prepare PDF Files:
   - Place all your PDF files in the pdfdata directory.

 Run the Script:
   - Execute the script to process the PDFs and generate the Excel file.


```javascript
python extract_pdf_data.py

```
Output
   - The script will create an Excel file named output_data.xlsx in the project root directory containing the extracted data.

## 🚀 About Me
I am Coding Enthusiast 😎

## 💫 Thank You

