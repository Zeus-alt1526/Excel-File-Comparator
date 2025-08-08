# Excel File Comparator (GUI)

A Python GUI tool to compare two Excel files. It allows you to select matching columns and then highlights:

- ✅ Matching rows
- ❌ Cell-level differences (highlighted in red)
- 🔍 Missing values (rows present in one file but not in the other)

## Features

- GUI using Tkinter
- Column-wise comparison
- Match, difference, and missing row detection
- Output saved to a single Excel file with multiple sheets

## Output Excel Sheets

- `Matching Rows File1`
- `Matching Rows File2`
- `Differences` (highlighted)
- `Missing in File2`
- `Missing in File1`

## Installation

Install required packages:

```bash
pip install -r requirements.txt
```

## Usage

Run the GUI tool:

```bash
python main.py
```

Steps:

1. Select Excel File 1 and File 2
2. Load both files
3. Select columns to compare from each file
4. Click **Compare and Export**
5. The output file `excel_comparison_output.xlsx` will be created in the same folder

## Dependencies

- pandas
- openpyxl
- tkinter (comes preinstalled with Python)
