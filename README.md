# 💼 Business Expense Tracker

A Python script that automatically analyses raw business expense data, detects unusual transactions using statistical methods, and generates a formatted Excel report and PDF summary — fully automated, one command.

---

## 🔍 What It Does

- Cleans and standardises raw CSV expense data
- Summarises spend by **department** and **category** per month
- Calculates **budget vs actual variance** for each category
- Detects **anomalous transactions** using the 2-sigma (standard deviation) method
- Exports a formatted **Excel report** with two sheets — Summary and Anomalies
- Auto-generates a **PDF report** with charts and an anomaly table

---

## 📊 Output

| File | Description |
|------|-------------|
| `outputs/expense_report_DD-MM-YYYY.xlsx` | Excel with Summary + Anomalies sheets |
| `outputs/expense_report.pdf` | Full PDF report with charts and flagged transactions |
| `outputs/chart1.png` | Bar chart — monthly spend by department |
| `outputs/chart2.png` | Line chart — actual vs budget per month |

---

## 🛠️ Tech Stack

| Library | Purpose |
|---------|---------|
| `pandas` | Data cleaning, grouping, merging, pivot tables |
| `openpyxl` | Excel file creation and multi-sheet writing |
| `matplotlib` | Chart generation |
| `fpdf2` | PDF report generation |

---

## 📁 Project Structure

```
expense_tracker/
│
├── data/
│   └── expenses.csv          # Drop your CSV here
│
├── outputs/                  # All generated files appear here
│
├── main.py                   # Main script
├── requirements.txt          # Dependencies
└── README.md
```

---

## ▶️ How to Run

**1. Clone the repository**
```bash
git clone https://github.com/YOUR_USERNAME/expense-tracker.git
cd expense-tracker
```

**2. Install dependencies**
```bash
pip install -r requirements.txt
```

**3. Add your data**

Drop your CSV file into the `data/` folder. Make sure it has these columns:

```
Date, Department, Category, Description, Amount, Budget
```

Date format should be `DD-MM-YYYY`.

**4. Update the filename in main.py**
```python
df = pd.read_csv('data/your_file.csv')
```

**5. Run the script**
```bash
python main.py
```

All output files will appear in the `outputs/` folder automatically.

---

## 🔬 How Anomaly Detection Works

For each expense category, the script calculates the **mean** and **standard deviation** of spending across all months.

Any transaction where:

```
Amount > mean + 2 × std
```

is flagged as an anomaly. This is the standard **2-sigma rule** used in financial data analysis — transactions that are statistically unusual compared to normal spending patterns for that category.

---

## 📌 Note

This is a **script-based tool** — not a web app. It is designed to demonstrate Python data automation skills including data cleaning, statistical analysis, Excel automation, chart generation, and PDF report creation.

A Streamlit web interface is planned as a future upgrade.

---

## 👤 Author

**Shridhar**  
Python Developer · Data Automation  
[GitHub](https://github.com/YOUR_USERNAME) · [LinkedIn](https://linkedin.com/in/YOUR_PROFILE)
