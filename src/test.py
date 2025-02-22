import pandas as pd

table_info = []
# # Store extracted information
# table_info.append({
#     "file_path": 'test',
#     "sheet_name": 'test',
#     "table_name": 'test',
#     "start_row_number": int(1),
#     "end_row_number": int(1),
#     "start_col_number": int(1),
#     "end_col_number": int(1)
# })

# Handle case where no tables are found
if not table_info:
    print(f"No tables found in the provided Excel file: {table_info}")
    table_info.append({
        "file_path": None,
        "sheet_name": None,
        "table_name": None,
        "start_row_number": None,
        "end_row_number": None,
        "start_col_number": None,
        "end_col_number": None,
    })
else:
    print(f"okay: {table_info}")

# Convert list to DataFrame
table_details = pd.DataFrame(table_info)
print(table_details)