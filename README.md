# DocuTestPro

**DocuTestPro** is a Python tool for generating professional-grade reports (PDF and Word) from test data.

## Features
- Normalizes test results from CSV files.
- Generates a bar chart showing test status distribution.
- Creates reports with:
  - A detailed summary of test metrics.
  - An embedded chart.
  - A detailed table of results.

## Requirements
- Python 3.x
- Libraries:
  - `pandas`
  - `matplotlib`
  - `fpdf`
  - `python-docx`

## Usage
1. Place your test data CSV file in the `input/` folder.
2. Run the script:
   ```bash
   python docutestpro.py
3. The reports will be generated in the output/ folder.

## Example CSV Format

| Test Case           | Status   | Execution Time | Comments               |
|---------------------|----------|----------------|------------------------|
| Login Functionality | Passed   | 2.5s           | All validations passed |
| Sign-Up Form        | Failed   | 3.2s           | Field validation failed|
| Search Feature      | Skipped  | 0.0s           | Test skipped           |

## Output
- PDF: test_report_<timestamp>.pdf
- Word: test_report_<timestamp>.docx