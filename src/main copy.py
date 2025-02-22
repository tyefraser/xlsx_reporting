import os
import pandas as pd
from openpyxl import load_workbook
from logger_config import logger
from utils import validate_file
from update_xlsx_data import replace_table_data, get_excel_table_details

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
    logger.info("📂 Adding file information to load info")

    # Validate that `file_info` contains only one key
    validate_single_key(file_info)

    for file_name, data_info in file_info.items():
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
            raise ValueError(f"❌ Error: Unsupported file extension: {file_extension}")

    return files_to_load


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

    logger.info(f"✅ Files to Load: {files_to_load}")

    # TO DO: Implement actual file loading logic

    logger.info("📂 INPUT DATA LOAD COMPLETE\n")
    return files_to_load
