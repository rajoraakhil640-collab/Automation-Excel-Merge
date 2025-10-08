🧾 Employee-Salary Data Merger (VLOOKUP-style)

This project reads two Excel files — one containing employee details and another containing salary data — then merges them automatically (like Excel VLOOKUP).  
It produces both Excel and CSV outputs with clear highlighting for missing salary data.

---

## 📦 Features

✅ Reads data from `employees.xlsx` and `salaries.xlsx`  
✅ Merges files using **Employee_ID**  
✅ Adds a new column `Status`:
- `"Matched"` → when salary is found  
- `"Not Found"` → when salary is missing  

✅ Saves output as both `.xlsx` and `.csv`  
✅ Adds **date & timestamp** to output filenames  
✅ Highlights “Not Found” rows in **red** using `openpyxl`  
✅ Handles errors gracefully (missing files, bad columns, etc.)  
✅ Supports **command-line interface (CLI)** for file input  

---

## ⚙️ Installation Instructions

### Step 1 — Install Python
Make sure you have **Python 3.14+** installed.  
You can verify with:
```bash
python --version
```

If not installed, download it from:  
🔗 https://www.python.org/downloads/

---

### Step 2 — Install Required Libraries

In your terminal or PyCharm terminal, run:

```bash
pip install pandas 
```

---

### Step 3 — Clone or Download the Project

Create a new folder (e.g. `ExcelMergeProject`) and add these files:
```
Automation_Assignment.py
employees.xlsx
salaries.xlsx
```

You can use the provided sample Excel files (with 100+ records).

---

---

## 🧪 Sample Test Data

### employees.xlsx
| Employee_ID | Name          | Department  |
|--------------|---------------|-------------|
| EMP001       | John Doe      | HR          |
| EMP002       | Alice Johnson | IT          |
| EMP003       | Mark Smith    | Sales       |

### salaries.xlsx
| Employee_ID | Salary |
|--------------|--------|
| EMP001       | 60000  |
| EMP003       | 72000  |

### Example Output Summary
```
=== MERGE SUMMARY ===
Total Employees : 3
Matched Records : 2
Unmatched Records : 1

Output saved as:
   Excel → merged_output_20251007_193015.xlsx
   CSV   → merged_output_20251007_193015.csv

Unmatched rows highlighted in red inside merged_output_20251007_193015.xlsx
```

## ▶️ How to Run the Project Locally

### Using PyCharm 

1. Open the folder in PyCharm
2. Open Automation_Assignment.py
3. Click the green ▶ Run button
4. If you get a prompt for arguments, enter:
   ```
   employees.xlsx salaries.xlsx -o merged_output

 Excel Output Example:

| Employee_ID | Name      | Department | Salary  | Status      |
|--------------|-----------|-------------|----------|--------------|
| EMP001       | Alice Lee | IT          | 60000    | Matched      |
| EMP002       | John Doe  | HR          | *blank*  | Not Found  |

Rows with missing salary values are highlighted in red*in Excel.


## 👨‍💻 Author
Akhil Rajora
📍 Data & Python Automation Enthusiast  
🗓️ Created: October 2025  


## 📜 License
This project is open-source and free to use for learning purposes
