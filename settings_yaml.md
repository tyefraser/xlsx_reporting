# 📖 Advanced Configuration for `settings.yaml`

### **Automating Data Processing & Excel Report Generation**

🚀 **Version:** 1.0.0 | 🛠 **Developed by:** Tye Fraser

---

## **📖 Overview**

This guide provides a detailed breakdown of how to configure the `settings.yaml` file to map **input data sources** (CSV/XLSX) to **output Excel reports**.

- 🔹 **Input Sources:** CSV files, Excel sheets, or Excel tables.
- 🔹 **Output Reports:** Excel templates structured with tables or sheets.
- 🔹 **Data Mapping:** Column mappings and data types.

---

## **📌 Understanding the Output Report Structure**

Once the **source data** is prepared, you can generate the **output report**. The report can be any sheet(s) in an **Excel template** and should pull data from the source sheets.

✅ **Best Practice:**  
- Use **tables** to structure the **output report**.
- Utilize **table formula references** to automatically populate the reports.

---

## **📌 Preparing Input Files**

Before linking input files, determine the required **input data structure**.

| Format  | Recommended? | Best Practice |
|---------|-------------|--------------|
| **CSV**  | ✅ Yes | Simplifies ingestion and avoids Excel formatting issues |
| **XLSX (Tables)** | ✅ Yes | Preferred for structured input data |
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

---

## **📌 Mapping Source Data in `settings.yaml`**

The **source data** can come from:
✅ **A CSV file**  
✅ **A sheet in an Excel file (`xlsx`)**  
✅ **A table from an Excel file (`xlsx`)**  

### **CSV Example**
```yaml
customer_data.csv:
  column_mapping:
    cust_id: Customer ID
    cust_name: Name
  column_types:
    cust_id: int
    cust_name: str
```

### **Sheet in an Excel File (`xlsx`)**
```yaml
employee_data.xlsx:
  xl_sheet:
    name: EmployeeSheet
    column_mapping:
      emp_id: Employee ID
      emp_salary: Salary
    column_types:
      emp_id: int
      emp_salary: float
```

### **Table from an Excel File (`xlsx`)**
```yaml
employee_data.xlsx:
  xl_table:
    name: EmployeesTable
    column_mapping:
      emp_id: Employee ID
      emp_role: Role
    column_types:
      emp_id: int
      emp_role: str
```

---

## **📌 Column Mapping & Data Types**

| **Key**         | **Description**                                        | **Example**      |
|----------------|------------------------------------------------|---------------|
| `column_mapping` | Maps **input column names → output column names** | `emp_id → Employee ID` |
| `column_types` | Defines the **expected data type** for each column | `emp_salary: float` |

---

## **📌 Example Input Data**

### **CSV Input File (`customer_data.csv`)**
| cust_id | cust_name  |
|---------|-----------|
| 1001    | John Doe  |
| 1002    | Jane Smith  |
| 1003    | Alice Brown  |

---

### **Excel Sheet Input (`employee_data.xlsx`)**
#### **EmployeeSheet**
| emp_id | emp_salary |
|--------|-----------|
| 2001   | 55000.00  |
| 2002   | 67000.50  |
| 2003   | 72000.00  |

---

### **Excel Table Input (`EmployeesTable` in `employee_data.xlsx`)**
| emp_id | emp_role   |
|--------|-----------|
| 3001   | Manager   |
| 3002   | Engineer  |
| 3003   | Analyst   |

---

## ✅ **Final Step: Validating Your Configuration**

Once your `settings.yaml` is complete, verify that:
- **All required inputs** are correctly defined.
- **Column names and data types** are correctly mapped.
- **No overlapping tables** exist in Excel templates.

---

### 🚀 **Now, you’re ready to automate Excel reports seamlessly!** 🎉
