# 📊 XLSX Reporting Automation

### Automated Data Processing & Report Generation for Excel Files

🚀 **Version:** 1.0.0 | 🛠 **Developed by:** Tye Fraser

---

## 📖 Overview

The **XLSX Reporting Automation** project is a Python-based tool designed to **automate data processing and report generation in Excel files**. This tool reads structured data from CSV/XLSX file(s), uploads this data to specified locations in existing Excel template(s), and saves the updated output report(s).

Built using **pandas** and **openpyxl**, this solution allows for **seamless data integration**, ensuring that Excel reports are consistently updated without manual effort.

---

## 📂 Repository Structure

```
xlsx_reporting/
│── inputs/                 # Contains configuration files and input datasets
│   ├── input_files/        # Raw CSV/XLSX data files
│   ├── xlsx_templates/     # Excel report templates
│   ├── settings.yaml       # Configuration file
│
│── logs/                   # Stores application logs
│
│── src/                    # Main source code directory
│   ├── load_config.py      # Configuration loader
│   ├── load_input_data.py  # Input file processing
│   ├── logger_config.py    # Logging setup
│   ├── main.py             # Main script (entry point)
│   ├── update_xlsx_data.py # Excel processing logic
│   ├── utils.py            # Utility functions
│
│── .gitignore              # Ignore unnecessary files
│── README.md               # Documentation file
│── requirements.txt        # Python dependencies
```

---

## 🚀 Features

✔ **Automated Excel Report Updates** – Reads, processes, and updates Excel reports dynamically.  
✔ **Supports Multiple Input Formats** – Works with **CSV** and **XLSX** files.  
✔ **Template-Based Processing** – Updates predefined Excel templates.  
✔ **Comprehensive Logging** – Logs each step of the process for easy debugging.  
✔ **Error Handling & Validation** – Ensures input files and templates are valid before processing.  
✔ **Command-Line Interface (CLI)** – Users can specify input/output folders via CLI arguments.

---

## 🛠 Setup & Installation

### 1️⃣ Prerequisites

- Python **3.8+**
- Virtual environment (recommended)
- Required Python packages (see `requirements.txt`)

### 2️⃣ Clone the Repository

```bash
git clone https://github.com/tyefraser/xlsx_reporting.git
cd xlsx_reporting
```

### 3️⃣ Create a Virtual Environment

```bash
# macOS/Linux
python -m venv venv
source venv/bin/activate

# Windows
python -m venv venv
venv\Scripts\activate
```

### 4️⃣ Install Dependencies

```bash
pip install -r requirements.txt
```

### 5️⃣ Update Report Templates, Input Files, and `settings.yaml`

#### 📌 Report Templates

It is often easiest to start with the **desired output `.xlsx` file**. This file will have **sheets for the input data** and **other sheets that rely on that data as output reports**. You can use **Excel `tables` or `sheets` as data sources**.

- **Tables**: It is recommended to use the `table` feature in Excel and create named tables for structured data input. Ensure **tables do not overlap on any rows**.
- **Sheets**: You can use data from the Excel sheets, but ensure only **one dataset per sheet** and that data starts in **cell A1**.

✅ **Recommended Practice:** Only include necessary columns of data to keep file size small.

#### 📌 Input Files

Once the **report template structure** is finalized, prepare the **input data files**:

- Store data in **CSV files** for better compatibility (recommended).
- Alternatively, **Excel sheets or tables** can also be used.

#### 📌 Configuring `settings.yaml`

This file maps **input data sources** to **output templates**.

1️⃣ **Define output report templates** under `output_from_input_dict`.  
2️⃣ **Specify source locations** for input data (`tables` or `sheets`).  
3️⃣ **List all column mappings** and **data types**.

**Example Structure for `settings.yaml`**

```yaml
output_from_input_dict:
  employee_report.xlsx:
    tables:
      employees_list:
        csv: employee_data.csv
        column_mapping:
          emp_id: Employee ID
          emp_name: Name
        column_types:
          emp_id: int
          emp_name: str
```

Please refer to the settings_yaml.md file for a more detailed explanation on how to populate the settings.yaml file.

---

## ⚡ Usage

### **Command-Line Execution**

Run the script with:

```bash
python src/main.py -i "inputs/input_files" -x "inputs/xlsx_templates"
```

✅ **Options:**
| Argument | Description | Default |
|----------|------------|---------|
| `-i, --input_files_folder` | Path to the folder containing input data | `inputs/input_files` |
| `-x, --xlsx_templates_folder` | Path to the folder containing Excel templates | `inputs/xlsx_templates` |
| `-o, --outputs_folder` | Path to the folder for the outputs to be copied to | `outputs` |
| `-d, --report_date` | Date the report is generated (YYYY-MM-DD) | System run date |
| `-c, --config_path` | Path to the config YAML file | `inputs/settings.yaml` |

---

## 📝 Logging & Error Handling

All logs are saved in the `logs/` folder, making it easy to trace errors and debugging messages.

**Common Errors & Fixes:**
| Error Message | Cause | Solution |
|--------------|-------|----------|
| `FileNotFoundError: Input folder does not exist` | The specified input path is incorrect | Check folder path and `settings.yaml` |
| `ValueError: Tables are overlapping` | Data tables in Excel overlap | Ensure tables have distinct row ranges |
| `ModuleNotFoundError: No module named 'pandas'` | Dependencies missing | Run `pip install -r requirements.txt` |

---

## 📌 Future Enhancements

🔹 **Support for Google Sheets Integration**  
🔹 **Database Connectivity for Input Data**  
🔹 **GUI for Non-Technical Users**

---

## 💡 Contributing

Want to improve this project? Feel free to submit a **pull request** or report issues! 😊

---

## 📜 License

This project is licensed under the **MIT License**. See `LICENSE` for details.

---

### 🚀 **Now you're all set! Run the script and automate your Excel reports like a pro!** 🎉
