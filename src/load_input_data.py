import os
import pandas as pd
from openpyxl import load_workbook
from pathlib import Path
from logger_config import logger
# from utils import validate_file
# from update_xlsx_data import replace_table_data, get_excel_table_details

# Constants
XL_TABLE = "xl_table"
XL_SHEET = "xl_sheet"

def validate_single_key(py_dict):
    """
    Validates that a dictionary contains only a single key.
    Logs an error and raises ValueError if multiple keys are present.

    Parameters:
        py_dict (dict): Dictionary to validate.

    Returns:
        bool: True if valid, otherwise raises ValueError.
    """
    if len(py_dict.keys()) != 1:
        error_msg = f"❌ Error: Expected a single key, but multiple were found: {list(py_dict.keys())}"
        logger.error(error_msg)
        raise ValueError(error_msg)
    return True


def add_file_to_load_info(file_info, files_to_load):
    """
    Updates the `files_to_load` dictionary with required files and columns.

    Ensures:
    - The dictionary entry exists before updating.
    - Only unique column names are stored.

    Parameters:
        file_info (dict): Dictionary containing file metadata (name, column mappings, types).
        files_to_load (dict): Dictionary tracking which files need to be processed.

    Returns:
        dict: Updated `files_to_load` dictionary.
    """

    # Validate that `file_info` contains only one key
    validate_single_key(file_info)

    for file_name, data_info in file_info.items():
        print(f"file_name:{file_name}")
        print(f"data_info:{data_info}")
        file_extension = os.path.splitext(file_name)[1].lower()  # Normalize file extension

        if file_extension == ".csv":
            # Initialize entry if not exists
            files_to_load.setdefault(file_name, {"cols": set()})

            # Extract column names from mappings and types
            column_mapping_keys = set(data_info.get("column_mapping", {}).keys())
            column_types_keys = set(data_info.get("column_types", {}).keys())

            # Update the set with unique column names
            files_to_load[file_name]["cols"].update(column_mapping_keys)
            files_to_load[file_name]["cols"].update(column_types_keys)

        elif file_extension == ".xlsx":
            # Initialize entry if not exists
            files_to_load.setdefault(file_name, {})

            for xl_type, xl_settings in data_info.items():
                print(f"xl_type:{xl_type}")
                print(f"xl_settings:{xl_settings}")
                if xl_type not in {XL_TABLE, XL_SHEET}:
                    raise ValueError(f"❌ Error: Unsupported `xl_type`: {xl_type}")

                category_key = "xl_tables" if xl_type == XL_TABLE else "xl_sheets"
                files_to_load[file_name].setdefault(category_key, {})

                xl_name = xl_settings.get("name")
                if not xl_name:
                    raise ValueError(f"❌ Error: Missing name for {xl_type}")

                files_to_load[file_name][category_key].setdefault(xl_name, {"cols": set()})

                # Extract column names from mappings and types
                column_mapping_keys = set(xl_settings.get("column_mapping", {}).keys())
                column_types_keys = set(xl_settings.get("column_types", {}).keys())

                # Update the set with unique column names
                files_to_load[file_name][category_key][xl_name]["cols"].update(column_mapping_keys)
                files_to_load[file_name][category_key][xl_name]["cols"].update(column_types_keys)

        else:
            err_msg = f"❌ Error: Unsupported file extension: {file_extension}"
            logger.error(err_msg)
            raise ValueError(err_msg)

    return files_to_load


def load_excel_sheet(file_path, sheet_name, columns_to_load):
    """
    Loads a specific sheet from an Excel file with selected columns.

    Parameters:
        file_path (str): Path to the Excel file.
        sheet_name (str): Name of the sheet to load.
        columns_to_load (list): List of column names to load.

    Returns:
        pd.DataFrame: DataFrame containing the specified columns from the sheet.

    Raises:
        FileNotFoundError: If the file does not exist.
        ValueError: If the sheet or columns are not found.
    """
    logger.info(f"file_path:{file_path}")
    logger.info(f"sheet_name:{sheet_name}")
    logger.info(f"columns_to_load:{columns_to_load}")

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"❌ Error: File '{file_path}' not found.")

    # Load the sheet with only the required columns
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=columns_to_load, engine="openpyxl")
    except ValueError as e:
        raise ValueError(f"❌ Error: {e}")

    return df


def load_input_data(input_files_folder, input_data_dict):

    for file_name, data_config in input_data_dict.items():
        file_extension = os.path.splitext(file_name)[1].lower()  # Normalize file extension

        if file_extension == '.csv':
            logger.info(f"TO DO - import csv data")

        elif file_extension == '.xlsx':
            file_path=os.path.join(input_files_folder, file_name)
            file_path=Path(file_path)
            print(f"file_path:{file_path}")
            wb = load_workbook(file_path, data_only=False)

            for xl_type, xl_config in data_config.items():
                if xl_type == "xl_sheets":
                    for sheet_name, cols_dict in xl_config.items():
                        data_pd = load_excel_sheet(
                            file_path=file_path,
                            sheet_name=sheet_name,
                            columns_to_load=list(cols_dict['cols'])
                        )
                        input_data_dict[file_name][xl_type][sheet_name]["data"] = data_pd

                elif xl_type == "xl_tables":
                    print(xl_config)
                    for table_name, cols_dict in xl_config.items():
                        print(f"table_name:{table_name}")
                        print(f"cols_dict:{cols_dict}")
                        for sheet in wb.worksheets:
                            print(f"sheet:{sheet}")
                            if hasattr(sheet, "tables"):  # Tables exist in the sheet
                                for table in sheet.tables.values():
                                    print(f"table:{table}")
                                    if table.name == table_name:
                                        if table.ref: # Returns the table range like "A1:C10"
                                            print(f"table.ref:{table.ref}")
                                            df = pd.DataFrame(sheet[table.ref])

                                            # Convert openpyxl Cell objects to values
                                            df = df.applymap(lambda cell: cell.value)
                                            
                                            # Set column headers
                                            df.columns = df.iloc[0]  # First row as header
                                            df = df[1:].reset_index(drop=True)  # Remove header row from data

                                            input_data_dict[file_name][xl_type][table_name]["data"] = df

    return input_data_dict



# # Loop through each table in xl_config
# for table_name in xl_config.keys():
#     for sheet in wb.worksheets:
#         table_range = get_table_range(sheet, table_name)
        
#         if table_range:
#             # Read the table range into a DataFrame
#             df = pd.DataFrame(sheet[table_range])
            
#             # Convert openpyxl Cell objects to values
#             df = df.applymap(lambda cell: cell.value)
            
#             # Set column headers
#             df.columns = df.iloc[0]  # First row as header
#             df = df[1:].reset_index(drop=True)  # Remove header row from data
            
#             # Store in dictionary
#             dataframes[table_name] = df
#             break  # Stop searching once the table is found


#                     print(xl_config)
#                     exit()


#                 else:
#                     err_msg = f"❌ Error: Unsupported xl_type: {xl_type}"
#                     logger.error(err_msg)
#                     raise ValueError(err_msg)

#         else:
#             raise ValueError(f"❌ Error: Unsupported file extension: {file_extension}")

#     return input_data_dict


def input_data_loader(input_files_folder, config):
    """
    Loads input data based on the configuration file.

    - Determines which files and columns need to be loaded.
    - Updates `files_to_load` dictionary to track processing requirements.

    Parameters:
        input_files_folder (str): Path to the folder containing input files.
        config (dict): Configuration dictionary specifying data sources.

    Returns:
        dict: Dictionary containing input data.
    """
    for _ in range(2): logger.info("")
    logger.info("-" * 50)
    logger.info("🚀 RUNNING INPUT DATA LOADER")
    logger.info("-" * 50)

    files_to_load = {}

    # Iterate through config to determine required input files
    for output_file, outputs in config.get("output_from_input_dict", {}).items():
        logger.info(f"📁 Processing Output File: {output_file}")

        # Process tables
        logger.info("-" * 50)
        logger.info("Processing tables input data")
        for _, file_info in outputs.get("tables", {}).items():
            files_to_load = add_file_to_load_info(file_info, files_to_load)

        # Process sheets
        logger.info("-" * 50)
        logger.info("Processing sheet input data")
        for _, file_info in outputs.get("sheets", {}).items():
            files_to_load = add_file_to_load_info(file_info, files_to_load)

    # Load file data
    input_data_dict = load_input_data(
        input_files_folder=input_files_folder,
        input_data_dict=files_to_load
    )

    logger.info("📂 INPUT DATA LOAD COMPLETE")

    return input_data_dict
