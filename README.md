# Bulk Certificate Generator

A simple JavaScript-based web application that generates personalized PDF certificates in bulk using data from an Excel file.

## Features

- Upload Excel files (`.xlsx` or `.xls`)
- Read recipient details directly in the browser
- Generate personalized certificates for each record
- Download certificates as PDF files
- No backend or database required
- Excel data is processed locally in the browser

## Excel Format

Your Excel file should contain the following columns:

| Name | Age | Amount |
|------|-----|--------|
| Rick | 20 | 5000 |
| Ethan | 21 | 2500 |
| Sam | 22 | 10000 |

## How It Works

1. Upload an Excel file containing `Name`, `Age`, and `Amount` columns.
2. Click **Generate Certificates**.
3. The application reads each record from the Excel file.
4. A personalized certificate is created for every valid record.
5. Each certificate is converted into a PDF and downloaded.

## Technologies Used

- HTML
- CSS
- JavaScript
- [SheetJS (XLSX)](https://sheetjs.com/) - Reading Excel files
- [html2canvas](https://html2canvas.hertzen.com/) - Converting certificates into images
- [jsPDF](https://github.com/parallax/jsPDF) - Generating PDF files

## Project Structure

```text
Bulk-Certificate-Generator/
│
└── index.html
└── index.css
└── index.js

