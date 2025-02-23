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

Detailed explanation to populate the yaml file:
Once you have included the source data, you can generate the outputs report. This can be any other sheet(s) in the xlsx file. This be structured in any way and pull data from the source sheets. Again, it is recommended to use `tables` to structure the output report and table formula references to source the required data to populate that report.

#### Input Files

Once you have determined the final structure for the report template, you should now know what input data is required for that report to be generated. you will then need to determine how to structure the input data files. It is recommended to store the data in .csv files for simplified data ingestion, however you can also use xlsx sheets and tables as input too. Please ensure you have all input sources listed and example files generated.

#### Settings.yaml

Now you will have your complete list of report output files and input files required to generate your reports. You should clear the report template, leaving all header rows, and for data tables leave one row of blank data. Now you can link each report template to the input data by populating the `settings.yaml` file. You will do this in the following stages:

- Determine the name of the output report template (e.g `employee_report.xlsx`) and add this to the `output_from_input_dict` section of the yaml file. In your output report template you will have input data sources being either `tables` or `sheets`
  - For `tables`: determine the name of the table in the output report (e.g `employees_list`) and add this under the `tables` key in the dict. To the named table (e.g. `employees_list`) add in the source data information (refer below on how to add source data information).
  - For `sheets`: determine the name of the sheet in the output report (e.g `hours`) and add this under the `sheets` key in the dict. To the sheet (e.g. `hours`) add in the source data information (refer below on how to add source data information).

Adding soure data information
As noted above you will need to add in the source data information for each of the report template source locations. The source data can come from:

- A csv file
- A sheet in an xlsx
- A table from an xlsx

CSV:
For a CSV file you will need to include the following dictionary keys:

- csv name: E.g. `customer_data.csv`. Under this you will need to include two sets of data `column_mapping` and `column_types`. Refer below on how this data is structured.

sheet in an xlsx:
For a sheet in an xlsx you will need to add the following dictionary keys and values:

- xlsx name: E.g `employee_data.xlsx`. To this you will need to add the key `xl_sheet` that indicates the data comes from an Excel sheet. In this key you need to add further keys for:
  - `name`: with the value being the name of the sheet
  - `column_mapping`: Refer below on how to structure this
  - `column_types`: Refer below on how to structure this

table from an xlsx:

- xlsx name: E.g `employee_data.xlsx`. To this you will need to add the key `xl_table` that indicates the data comes from an Excel table. In this key you need to add further keys for:
  - `name`: with the value being the name of the table in the xlsx file
  - `column_mapping`: Refer below on how to structure this
  - `column_types`: Refer below on how to structure this

For each input table, you will need to include dictionaries for `column_mapping`, and `column_types`:

- `column_mapping`: This will include a dictionary with keys for all input column headers and values for all of the corresponding header names in the output template file.
- `column_types`: This will include a dictionary with keys for all input column headers and values for all of the data types for those columns.

You should now have a yaml file that details all of the required output templates and how to populate the required input data from soruce files.

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
