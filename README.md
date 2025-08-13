# Excel File Comparator

A Python Tkinter-based GUI tool to compare two Excel files.  
It can:
- **Find Differences** between rows based on a selected key column.
- **Find Missing Data** in either file.
- Export results to a new Excel file with highlighted changes.

---

## 📦 Features
- Load and view columns from two `.xlsx` / `.xls` files.
- Select the key column in each file for comparison.
- Choose between:
  - **Find Differences** → Compares matching rows cell-by-cell.
  - **Find Missing Data** → Finds rows missing in either file.
- Automatically highlights differences in the output Excel file.
- Prevents overwriting results by auto-generating unique filenames.

---

## 🚀 Installation

### 1. Clone the Repository
```bash
git clone https://github.com/YOUR_USERNAME/excel-file-comparator.git
cd excel-file-comparator
```

### 2. Create Virtual Environment (Recommended)
```bash
python -m venv venv
# Activate:
# Windows:
venv\Scripts\activate
# Mac/Linux:
source venv/bin/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```
> **Note:** `tkinter` comes pre-installed with most Python distributions.  
> On Ubuntu/Debian you may need: `sudo apt install python3-tk`  
> On macOS with Homebrew: `brew install python-tk`

---

## 🖥 Usage
```bash
python excel_comparator.py
```
1. Click **Browse** to select each Excel file.
2. Click **Load File 1** and **Load File 2** to load them.
3. Select the key column in each file.
4. Choose the action:
   - `Find Differences`
   - `Find Missing Data`
5. Click **Compare and Export** — results will be saved to an Excel file in the same folder.

---

## 📂 Project Structure
```
excel-file-comparator/
│── excel_comparator.py    # Main application script
│── requirements.txt       # Python dependencies
│── README.md              # Documentation
└── .gitignore             # Ignore unnecessary files
```

---

## 🛠 Dependencies
- Python 3.8+
- pandas
- openpyxl
- tkinter (bundled with Python)

Install with:
```bash
pip install pandas openpyxl
```

---

## 📸 Screenshots
*(Add screenshots of your GUI here)*

---

## 📜 License
This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.
