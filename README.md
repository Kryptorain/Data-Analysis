<img src="lol.jpg" />


# 📊 Data Analysis using Excel with Python

Welcome to **Kryptora's** Excel-powered data analysis project, developed by **Manish (@manishhacker1)**. This notebook demonstrates how to create, manipulate, and export structured data to Excel using `pandas` in Python.

---

## 📁 Project Structure

This project includes:
- Creating sample DataFrames
- Exporting to a **single Excel sheet**
- Writing multiple **DataFrames to different Excel sheets**
- Demonstrating **Excel file automation** using `pandas` and `openpyxl`

---

## 📌 Key Features

✔️ Create DataFrames manually  
✔️ Export data to `.xlsx` files  
✔️ Use `ExcelWriter` to manage multiple sheets  
✔️ Real-world tabular data simulation for city-wise and person-wise data

---

## 💻 Sample Code Snippet

```python
import pandas as pd

data = {
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [25, 30, 22],
    'City': ['Delhi', 'Mumbai', 'Bangalore']
}
df = pd.DataFrame(data)
df.to_excel("sample_data.xlsx", index=False)
```

---

## 📂 Output Files

- `sample_data.xlsx` → One-sheet Excel file  
- `multi_sheet_data.xlsx` → Multi-sheet Excel file (`People`, `Cities`)

---

## 📚 Prerequisites

Make sure you have the following installed:
```bash
pip install pandas openpyxl
```

---

## 🚀 How to Run

```bash
jupyter notebook Excel.ipynb
```

---

## 🤝 Contributing

Pull requests are welcome! For major changes, please open an issue first to discuss.

---

## 👨‍💻 About the Author

**Manish** (aka [@manishhacker1](https://github.com/manishhacker1)) is a data analyst and founder of **Kryptora**, building educational projects in Python, Excel, and data science.
