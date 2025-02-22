import os
import shutil

def delete_file(file_path):
    """Deletes a file if it exists."""
    if os.path.exists(file_path):
        os.remove(file_path)
        print(f"Deleted: {file_path}")
    else:
        print(f"File not found: {file_path}")

def copy_and_rename_file(source_path, destination_path):
    """Copies and renames a file."""
    if os.path.exists(source_path):
        shutil.copy(source_path, destination_path)
        print(f"Copied and renamed: {source_path} → {destination_path}")
    else:
        print(f"Source file not found: {source_path}")

# Usage
file_to_delete = r"C:\Users\tyewf\github_projects\xlsx_reporting\inputs\xlsx_templates\template_1.xlsx"
source_file = r"C:\Users\tyewf\github_projects\xlsx_reporting\inputs\xlsx_templates\template_1 - Original.xlsx"
destination_file = file_to_delete
delete_file(file_to_delete) # Delete a file
copy_and_rename_file(source_file, destination_file) # Copy and rename a file

# Usage
file_to_delete = r"C:\Users\tyewf\github_projects\xlsx_reporting\inputs\xlsx_templates\employee_report.xlsx"
source_file = r"C:\Users\tyewf\github_projects\xlsx_reporting\inputs\xlsx_templates\employee_report - Original.xlsx"
destination_file = file_to_delete
delete_file(file_to_delete) # Delete a file
copy_and_rename_file(source_file, destination_file) # Copy and rename a file

# Usage
file_to_delete = r"C:\Users\tyewf\github_projects\xlsx_reporting\inputs\xlsx_templates\industry_report.xlsx"
source_file = r"C:\Users\tyewf\github_projects\xlsx_reporting\inputs\xlsx_templates\industry_report - Original.xlsx"
destination_file = file_to_delete
delete_file(file_to_delete) # Delete a file
copy_and_rename_file(source_file, destination_file) # Copy and rename a file

# Usage
file_to_delete = r"C:\Users\tyewf\github_projects\xlsx_reporting\inputs\xlsx_templates\customer_report.xlsx"
source_file = r"C:\Users\tyewf\github_projects\xlsx_reporting\inputs\xlsx_templates\customer_report - Original.xlsx"
destination_file = file_to_delete
delete_file(file_to_delete) # Delete a file
copy_and_rename_file(source_file, destination_file) # Copy and rename a file