<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PDF Data Extractor and Excel Writer</title>
  <style>
    body {
      font-family: Arial, sans-serif;
      line-height: 1.6;
      background-color: #f4f4f4;
      padding: 20px;
    }

    h1 {
      color: #333;
    }

    p {
      color: #555;
    }

    pre {
      background-color: #eee;
      padding: 10px;
      border-radius: 5px;
      overflow-x: auto;
    }

    table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 20px;
    }

    th, td {
      padding: 10px;
      border-bottom: 1px solid #ddd;
      text-align: left;
    }

    th {
      background-color: #f2f2f2;
    }

    .container {
      max-width: 800px;
      margin: auto;
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>PDF Data Extractor and Excel Writer</h1>

    <p>This repository contains a Python-based solution designed to extract specific data from multiple PDF files and compile the extracted information into an Excel file. Utilizing the PyMuPDF library for PDF processing and the openpyxl library for Excel file handling, this project automates the extraction and organization of data from a folder of PDF files into a structured Excel spreadsheet.</p>

    <h2>Features</h2>
    <ul>
      <li><strong>Automated PDF Data Extraction</strong>: Extracts designated fields from PDFs using regular expressions.</li>
      <li><strong>Data Structuring</strong>: Converts location data into separate columns for district, taluka, and village.</li>
      <li><strong>Excel Output</strong>: Saves the extracted data into an Excel file with clear column headings.</li>
      <li><strong>Row Indexing</strong>: Adds an index to each row in the Excel file for easy identification.</li>
    </ul>

    <h2>Prerequisites</h2>
    <ul>
      <li>Python 3.6 or higher</li>
      <li>Required Python libraries:
        <ul>
          <li>PyMuPDF (`fitz`)</li>
          <li>openpyxl</li>
        </ul>
      </li>
    </ul>

    <h2>Installation</h2>
    <pre>
      <code>git clone https://github.com/yourusername/pdf-data-extractor.git
cd pdf-data-extractor
pip install pymupdf openpyxl</code>
    </pre>

    <h2>Project Structure</h2>
    <pre>
      <code>pdf-data-extractor/
  ├── pdfdata/
  │   ├── file1.pdf
  │   ├── file2.pdf
  │   └── file3.pdf
  ├── extract_pdf_data.py
  └── README.md</code>
    </pre>

    <p>Replace `yourusername` with your GitHub username and `[your email]` with your contact email.</p>
  </div>
</body>
</html>
