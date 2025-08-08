# Excel File Comparator (GUI)

A Python GUI tool to compare two Excel files. It allows you to select matching columns and then highlights:

- ✅ Matching rows
- ❌ Cell-level differences (highlighted in red)
- 🔍 Missing values (rows present in one file but not the other)

## 💻 Features

- GUI using Tkinter
- Column-wise comparison
- Match, difference, and missing row detection
- Output saved to a single Excel file with multiple sheets

## 📂 Output File Includes:

- `Matching Rows File1`
- `Matching Rows File2`
- `Differences` (highlighted)
- `Missing in File2`
- `Missing in File1`

## 🛠 Installation

Install required packages:

```bash
pip install -r requirements.txt
