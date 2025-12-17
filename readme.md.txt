# Contract Drafting Web App

This is a FastAPI-based web application that generates Word (.docx)
contracts from Excel input files using predefined DOCX templates.

## How it works
1. Upload a DOCX template with {{placeholders}}
2. Upload an Excel file (headers = placeholders)
3. Download generated contracts as a ZIP file

## Tech Stack
- FastAPI
- Pandas
- docxtpl

This tool is for internal/educational use.
