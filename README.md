# FAMS – Form Automation Management System

FAMS (Form Automation Management System) is a Python-based desktop application that automates the generation of personalized documents for students using a Word template and a CSV or Excel file.

It is designed for academic institutions to quickly generate forms, letters, certificates, or notices for large student lists with minimal effort.

---

## ✨ Features

- 📂 Import student data from **CSV / Excel (XLS, XLSX)**
- 📝 Use a **DOCX template** with placeholders
- ⚡ Generate **individual DOCX files** per student
- 📄 Optional **PDF generation**
- 📚 Merge all DOCX files into one document
- 🧾 Merge all PDFs into one file
- 📊 Real-time progress bar
- 🖥 GUI built with Tkinter & CustomTkinter
- 🧾 Activity logging with downloadable logs
- ❓ Built-in Help / User Guide

---

## 🖼 Application Overview

- Splash screen on startup  
- Modern and user-friendly interface  
- Live activity logs and progress tracking  
- Help window with screenshots and tips  

---

## 📁 Project Structure
```
├── main.py
├── assets/
|   ├── mbc.ico
│   ├── mbc.png
│   ├── splash.png
│   ├── browse.png
│   ├── check.png
│   ├── help.png
│   ├── genrate.png
│   ├── openfolder.png
│   ├── log.png
│   ├── clear.png
│   ├── ss_example.png
│   └── ss1_example.png
├── fams_output/
│ ├── docx/
│ ├── pdf/
│ ├── merged_docx/
│ ├── merged_pdf/
│ └── fams_log.txt
└── README.md
```

---

## 🧑‍🎓 Student File Format

Supported formats:
- CSV
- XLS
- XLSX

### Required Columns

The application automatically detects the following columns:

**Name**
- `name`

**Student Number**
- `student_number`

If column names are not detected, the **first two columns** will be used automatically.

---

## 📄 DOCX Template Placeholders

Use the following placeholders in your Word template:

```
{{ name }}
{{ student_number }}
```

---

## 📌 Template Formatting Rules

- One student generates **one document or one page**
- Use **manual page breaks** (`Ctrl + Enter`)
- Do **NOT** add extra blank pages at the end
- Place all placeholders on the same page
- Avoid placeholders inside:
  - Text boxes
  - Shapes
- Headers and footers are supported
- Page breaks control merged document layout

---

## 🚀 How to Use

1. Launch the application
2. Click **Browse** to upload student data (CSV / Excel)
3. Click **Browse** to select a DOCX template
4. Choose optional actions:
   - Generate PDF
   - Merge all DOCX
   - Merge all PDFs
5. Click **Generate Documents**
6. Monitor progress and logs
7. Open the output folder or download logs

---

## 📂 Output Directory

All generated files are saved in:
`fams_output/`

### Subfolders

- `docx/` – Individual Word documents
- `pdf/` – Individual PDF files
- `merged_docx/` – Combined DOCX file
- `merged_pdf/` – Combined PDF file
- `fams_log.txt` – Activity logs

---

## 🛠 Requirements

### Python Version
- Python **3.12+**

### Required Python Packages

```bash
pip install -r requirements.txt
```
### Run Program
```bash
python main.py
```
