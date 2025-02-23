# 📊 XLSX Reporting Automation

### Automated Data Processing & Report Generation for Excel Files

🚀 **Version:** 1.0.0 | 🛠 **Developed by:** [Your Name]

---

## 📖 Overview

The **XLSX Reporting Automation** project is a Python-based tool designed to **automate data processing and report generation in Excel files**. This tool reads structured data from CSV/XLSX files, updates existing Excel templates, and generates formatted reports efficiently.

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
git clone https://github.com/your-username/xlsx_reporting.git
cd xlsx_reporting
```

### 3️⃣ Create a Virtual Environment

```bash
python -m venv venv
source venv/bin/activate  # macOS/Linux
venv\Scripts\activate   # Windows
```

### 4️⃣ Install Dependencies

```bash
pip install -r requirements.txt
```

---

## ⚡ Usage

### Command-Line Execution

Run the script with:

```bash
python src/main.py -i "inputs/input_files" -x "inputs/xlsx_templates"
```

✅ **Options:**
| Argument | Description | Default |
|----------|------------|---------|
| `-i, --input_files_folder` | Path to the folder containing input data | `inputs/input_files` |
| `-x, --xlsx_templates_folder` | Path to the folder containing Excel templates | `inputs/xlsx_templates` |

---

## ⚙️ Configuration

The `settings.yaml` file allows you to define **custom configurations**, such as:

```yaml
input_files_folder: "inputs/input_files"
xlsx_templates_folder: "inputs/xlsx_templates"
log_level: "INFO"
```

---

## 📌 Example Workflow

1️⃣ **Prepare Data** – Place CSV/XLSX files into `inputs/input_files/`.  
2️⃣ **Define Templates** – Ensure your Excel templates exist in `inputs/xlsx_templates/`.  
3️⃣ **Run the Script** – Execute `python src/main.py`.  
4️⃣ **Generate Reports** – Updated reports will be stored in `outputs/`.

---

## 📝 Logging & Error Handling

All logs are saved in the `logs/` folder, making it easy to trace errors and debugging messages.

**Common Errors & Fixes:**
| Error Message | Cause | Solution |
|--------------|-------|----------|
| `FileNotFoundError: Input folder does not exist` | The specified input path is incorrect | Check folder path and settings.yaml |
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
