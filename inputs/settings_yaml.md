# 📖 Detailed instructions for `settings.yaml`

### **Automating Data Processing & Excel Report Generation**

🚀 **Version:** 1.0.0 | 🛠 **Developed by:** Tye Fraser

---

## **📖 Overview**

This guide provides a detailed breakdown of how to configure the `settings.yaml` file to map **input data sources** (CSV/XLSX) to **output Excel reports**.

- **Input Sources:** CSV files, Excel sheets, or Excel tables.
- **Output Reports:** Excel templates structured with tables or sheets.
- **Data Mapping:** Column mappings and data types.

---

## **📌 Understanding the Output Report Structure**

Once the **output report** is designed, you can determine the **source data** required. At this stage you should have an output report template xlsx file that uses data from named Excel tables and/or sheets filled with data (starting from A1 and only one dataset per sheet).

✅ **Best Practice:**

- Use **tables** to structure the **output report**.
- Utilise **table formula references** to automatically populate the reports.

---

## **📌 Preparing Input Files**

Before linking input files, determine the required **input data structure**.

| Format            | Recommended?              | Best Practice                                           |
| ----------------- | ------------------------- | ------------------------------------------------------- |
| **CSV**           | ✅ Yes                    | Simplifies ingestion and avoids Excel formatting issues |
| **XLSX (Tables)** | ✅ Yes                    | Preferred for structured input data                     |
| **XLSX (Sheets)** | ⚠️ Yes (Use with caution) | Ensure **one dataset per sheet**, data starts at **A1** |

✅ **Checklist Before Proceeding:**  
✔ Ensure all **input sources are listed** and **example files are available**.  
✔ Structure data **logically and consistently** for all input files.

---

## **📌 Configuring `settings.yaml`**

Now that your **output reports** and **input files** are finalized, you can define their relationships in **`settings.yaml`**.

### **📝 Steps to Populate `settings.yaml`**

1️⃣ Identify the **output report template** (e.g., `employee_report.xlsx`).  
2️⃣ Add the template to the **`output_from_input_dict`** section.  
3️⃣ Define the input **tables or sheets** inside the output template.  
4️⃣ Specify **source data information** (CSV, sheets, or tables).  
5️⃣ Provide **column mapping** and **column types** to ensure consistency.

---

## **📌 Structure of `settings.yaml`**

This file defines how input data is mapped into the Excel report.

### **📝 Example Output Template Definition**

```yaml
output_from_input_dict:
  customer_report.xlsx:
    tables:
      customers:
        customer_data.csv:
          column_mapping:
            "First": "First Name"
            "Last": "Last Name"
            "Start": "Start Date"
            "Sales": "Total Sales"
          column_types:
            "First": string
            "Last": string
            "Start": date
            "Sales": float64
```

---

## **📌 Mapping Source Data in `settings.yaml`**

The **source data** can come from:
✅ **A CSV file**  
✅ **A sheet in an Excel file (`xlsx`)**  
✅ **A table from an Excel file (`xlsx`)**

### **CSV Example**

```yaml
output_from_input_dict:
  customer_report.xlsx: # Name of the template report
    tables: # Specifies that data is to be added to the tables listed below to the template report
      customers: # name of the table in the template to add data to
        customer_data.csv: # source file from where the data comes from << This is the CSV example for sourcing input data>>
          column_mapping: # Maps the source columns to the report template columns
            "First": "First Name" # 'First' is the name of the input column, 'First Name' is the name of the corresponding column in the report template
            "Last": "Last Name"
            "Start": "Start Date"
            "Sales": "Total Sales"
          column_types: # Provides data types for the source data
            "First": string
            "Last": string
            "Start": date
            "Sales": float64
    sheets: # Specifies that the input data is to be added to the following sheets
      customers_flat: # Name of the sheet to add the data to
        customer_data.csv: # name of the source CSV file to obtain the data
          column_mapping:
            "First": "First Name"
            "Last": "Last Name"
            "Start": "Start Date"
            "Sales": "Total Sales"
          column_types:
            "First": string
            "Last": string
            "Start": date
            "Sales": float64
```

### **Sheet in an Excel File (`xlsx`)**

```yaml
output_from_input_dict:
  employee_report.xlsx: # Name of the template report
    tables: # Specifies the tables in the Excel template that you want to add data to
      employees_list: # Name of the table
        employee_data.xlsx: # Name of the xlsx to source data from
          xl_sheet: # Specifies that the data comes from the whole sheet ('xl_table' would specify a table to source the data from)
            name: "Employees" # name of the sheet in the input xlsx file
            column_mapping:
              "First": "First Name"
              "Last": "Last Name"
              "Id": "Identifier"
            column_types:
              "First": string
              "Last": string
              "Id": string
      employee_hours:
        employee_data.xlsx:
          xl_sheet:
            name: "Hours"
            column_mapping:
              "First": "First Name"
              "Last": "Last Name"
              "Rate": "Rate"
              "Hours": "Hours"
            column_types:
              "First": string
              "Last": string
              "Rate": float64
              "Hours": float64
```

### **Table from an Excel File (`xlsx`)**

```yaml
output_from_input_dict:
  industry_report.xlsx: # name of the template file
    tables: # Specifies the tables to add data to
      companies: # name of the table in the template xlsx
        industry_data.xlsx: # Name of the input file
          xl_table: # Specifies that the data comes from an Excel table
            name: "companies_input" # name of the table in the input excel file
            column_mapping:
              "Company Name": "Company Name"
              "Shares": "Shares Outstanding"
              "Price": "Share Price"
            column_types:
              "Company Name": string
              "Shares": float64
              "Price": float64
      region_stats: # name of the table in the template xlsx
        industry_data.xlsx: # Name of the input file
          xl_table: # Specifies that the data comes from an Excel table
            name: "regions" # name of the table in the input excel file
            column_mapping:
              "Regions": "Region Code"
              "Sales": "Total Sales"
            column_types:
              "Regions": string
              "Sales": float64
```

---

## **📌 Column Mapping & Data Types**

| **Key**          | **Description**                                    | **Example**             |
| ---------------- | -------------------------------------------------- | ----------------------- |
| `column_mapping` | Maps **input column names → output column names**  | `emp_id: 'Employee ID'` |
| `column_types`   | Defines the **expected data type** for each column | `emp_salary: float`     |

---

## **📌 Example Input Data**

### **CSV Input File (`customer_data.csv`)**

| cust_id | cust_name   |
| ------- | ----------- |
| 1001    | John Doe    |
| 1002    | Jane Smith  |
| 1003    | Alice Brown |

---

### **Excel Sheet Input (`employee_data.xlsx`)**

#### **EmployeeSheet**

| emp_id | emp_salary |
| ------ | ---------- |
| 2001   | 55000.00   |
| 2002   | 67000.50   |
| 2003   | 72000.00   |

---

### **Excel Table Input (`EmployeesTable` in `employee_data.xlsx`)**

| emp_id | emp_role |
| ------ | -------- |
| 3001   | Manager  |
| 3002   | Engineer |
| 3003   | Analyst  |

---

## ✅ **Final Step: Validating Your Configuration**

Once your `settings.yaml` is complete, verify that:

- **All required inputs** are correctly defined.
- **Column names and data types** are correctly mapped.
- **No overlapping tables** exist in Excel templates.

---

### 🚀 **Now, you’re ready to automate Excel reports seamlessly!** 🎉
