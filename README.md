# PathologyOCR
Automatically extract structured blood test results from scanned or faxed lab report images using Python and Tesseract OCR, then map and export them to a clean, analysis-ready spreadsheet.
**Doesn't Quite Work Yet!!**

## 🚀 Features

- **Robust OCR:** Handles messy sources (fax, scan, emailed PDFs).
- **Standardized Test Mapping:** Translates varied and error-prone test names into standardized labels.
- **Multi-Date Support:** Detects and aligns results from multiple test dates.
- **Excel Template Integration:** Fills your own Excel template for further review or analysis.
- **Debug Outputs:** Saves OCR text and processing logs for easy troubleshooting.

## 📋 Requirements

- Python 3.7+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) (install and note your install path)
- Python packages:
  - `pytesseract`
  - `pandas`
  - `Pillow`
  - `openpyxl`
 
## 💡 How It Works

1. **OCR Extraction:** Reads all table text from the image using Tesseract, tuned for tabular formats.
2. **Date Detection:** Locates and extracts all result date columns.
3. **Result Parsing:** For each row, standardizes the test name and captures corresponding results.
4. **Template Population:** Fills your Excel template, aligning results under the right columns.
5. **Export:** Outputs a structured CSV or Excel file for easy use.

## 📓 Example Workflow

**Input:**  
A scan or photo of a blood biochemistry report, such as `your_blood_test_image.jpg`.

**Excel Template:**  
Your template (e.g., `PathologyPro_Template.xlsx`) should have a "Blood Tests" sheet structured like:

| Test        | Units   | 15/04/20 | 03/03/22 | ... | Reference |
|-------------|---------|----------|----------|-----|-----------|
| Sodium      | mmol/L  |          |          |     | 135-145   |
| Potassium   | mmol/L  |          |          |     | 3.5-5.5   |
| ...         | ...     | ...      | ...      | ... | ...       |

**Output:**  
A populated CSV or Excel file, e.g., `populated_blood_test_results.csv`.

## 🛠 Usage

1. **Scan or photograph** your lab report and save it as an image.
2. **Prepare/validate your Excel template**.
3. **Update paths** in your script:
   - Image file (e.g., `your_blood_test_image.jpg`)
   - Excel template file (e.g., `PathologyPro_Template.xlsx`)
4. **Run the script:**

## 🧩 Troubleshooting

- **Did not extract data correctly?**
  - Check the saved `debug_ocr_output.txt` for OCR results and add mappings if needed.
- **Template mismatch?**
  - Ensure the Excel file and test names align with those expected by your parser.
- **Tesseract issues?**
  - Double-check Tesseract installation and set the correct path for your OS.

## 📜 License

MIT License

## 🙏 Acknowledgements

Developed to automate and simplify the extraction of health data from legacy documents—enabling better personal health recordkeeping, research, and interoperability. Contributions and suggestions are welcome!

**Let’s automate better health data—one messy PDF at a time.**

