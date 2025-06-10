# Report Generation Script

## Description
This script parses a `.docx` file containing structured entries and generates a report with formatted tables for each entry. Each entry includes fields like **Title, Date, Country (with flag), Summary, Link, and Availability**.

---

## How to Use on Replit

**Everything is already set up for you! Just follow these simple steps:**

1. **Upload your own `input.docx` file:**
   - On the left, find the file called `input.docx`.
   - Right-click it and choose **Delete** (or just overwrite it).
   - Drag and drop your own `.docx` file into the file list on the left, and make sure it is named `input.docx`.

2. **Click the green "Run" button at the top of the Replit window.**
   - By default, this will generate a **Word report** (`output.docx`).

3. **Download your result:**
   - When the script finishes, you will see a new file called `output.docx` in the file list.
   - Right-click `output.docx` and choose **Download** to save it to your computer.

---

## Need a PDF instead of a Word file?

- Open the `.replit` configuration file in the file list.
- On line 2, you will see the default run command:
  ```
  run = "python generate_reports.py input.docx -w"
  ```
- **Remove the `-w` part** so it looks like this:
  ```
  run = "python generate_reports.py input.docx"
  ```
- Click the green **Run** button again. This will generate `output.pdf` instead of `output.docx`.
- When the script finishes, download the final file (`output.pdf`) from the file list as above.

---

## Notes

- You do **not** need to install anything or use the Shell on Replit.
- If you get stuck, just ask for help!
- Make sure your input file is named **exactly** `input.docx` and is in the main file list (not inside a folder).
- The country flag images are already included for you.

---

## Running Locally on Unix/Linux/Mac

1. **Install Python 3.6 or newer** (if not already installed).

2. **Clone or download this repository** to your computer.

3. **(Recommended) Create and activate a virtual environment:**
   ```sh
   python3 -m venv venv
   source venv/bin/activate
   ```

4. **Install the required packages:**
   ```sh
   pip install -r requirements.txt
   ```

5. **Place your input file as `input.docx` in the project folder.**

6. **To generate a Word report:**
   ```sh
   python generate_reports.py -w
   ```

7. **To generate a PDF report:**
   ```sh
   python generate_reports.py
   ```

8. **Your output will be saved as `output.docx` or `output.pdf` in the same folder.**

---

## Running Unit Tests

1. Make sure you have installed all requirements (see above).

2. From the project root, run:
   ```sh
   python -m unittest discover -s test
   ```
   or to run a specific test file:
   ```sh
   python -m unittest test/test_generate_reports.py
   ```

---
