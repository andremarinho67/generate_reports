Report Generation Script
========================

Description:
------------
This script parses a .docx file containing structured entries and generates a report with formatted tables for each entry. Each entry includes fields like Title, Date, Country (with flag), Summary, Link, and Availability.

Requirements:
--------------
- Python 3.6+
- Packages listed in requirements.txt (install via pip)

Setup:
-------
1. Create a virtual environment to isolate dependencies (recommended):

   On macOS/Linux:
   python3 -m venv venv
   source venv/bin/activate

   On Windows:
   python -m venv venv
   venv\Scripts\activate

2. Install required packages:

   pip install -r requirements.txt

Note:
------
- The venv/ folder is not included in this repository and should not be committed to Git.
- Virtual environments are machine-specific. Cloning this repo requires creating your own venv and installing dependencies using the requirements.txt file.

Usage:
-------
- Place your input .docx file as input.docx in the working directory or adjust the script accordingly.
- Run the script to generate output.pdf with formatted entries by default.
- To generate a Word (.docx) report instead, use the -w flag.

Flags:
-------
- Flag images for countries should be placed in the flags/ directory as PNG files.
- The script automatically loads and displays them in the PDF or Word report.

Command examples:
-----------------
Generate PDF (default):
   python generate_reports.py

Generate Word document:
   python generate_reports.py -w

Troubleshooting:
----------------
- If you get a ModuleNotFoundError for missing packages, ensure your virtual environment is activated and dependencies installed.
- Use pip install -r requirements.txt to install missing packages.

---
