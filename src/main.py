import argparse
import os
import yaml
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

# Importing custom modules
from update_xlsx_data import replace_table_data, get_excel_table_details
from load_config import config_loader
from utils import validate_folder, validate_file, is_valid_date
from logger_config import logger
from load_input_data import input_data_loader
from update_xlsx_data import add_data_to_files

# 🔹 Parse Command-Line Arguments
def parse_args():
    """Parses command-line arguments for sourcing input/output folders and configurations."""
    for _ in range(2): logger.info("")
    logger.info("-" * 50)
    logger.info("🔍 Extracting and validating arguments...")
    logger.info("-" * 50)

    parser = argparse.ArgumentParser(
        description="Batch process Excel templates, updating tables with new data sources."
    )

    parser.add_argument(
        "-i", "--input_files_folder",
        default="inputs/input_files",
        help="Path to the `input_files` folder (default: inputs/input_files)"
    )
    parser.add_argument(
        "-x", "--xlsx_templates_folder",
        default="inputs/xlsx_templates",
        help="Path to the `xlsx_templates` folder (default: inputs/xlsx_templates)"
    )
    parser.add_argument(
        "-o", "--outputs_folder",
        default="outputs",
        help="Path to the `outputs` folder where outputs will be stored (default: outputs)"
    )
    today_date = datetime.today().strftime('%Y-%m-%d')
    parser.add_argument(
        "-d", "--report_date",
        default=today_date,
        help=f"Run date in format YYYY-MM-DD (default: {today_date})"
    )
    parser.add_argument(
        "-c", "--config_path",
        default="inputs/settings.yaml",
        help="Path to the `settings.yaml` configuration file (default: inputs/settings.yaml)"
    )

    logger.debug(f"Arguments loaded")
    args = parser.parse_args()

    # Assign to separate variables
    input_files_folder = args.input_files_folder
    xlsx_templates_folder = args.xlsx_templates_folder
    outputs_folder = args.outputs_folder
    report_date = args.report_date
    config_path = args.config_path

    # Perform necessary validations
    folder_list = [input_files_folder, xlsx_templates_folder]
    for folder in folder_list:
        logger.debug(f"Checking folder exists: {folder}")
        if not validate_folder(folder):
            raise FileNotFoundError(f"Missing folder: {folder}")
    if not validate_file(config_path):
        logger.debug(f"Checking config file exists: {config_path}")
        raise FileNotFoundError(f"Missing file: {config_path}")
    logger.debug(f"Checking report_date: {report_date}")
    is_valid_date(report_date)

    # Create Outputs folder if it doesn't exist
    if not os.path.exists(outputs_folder):
        logger.info(f"📁 Folder `{outputs_folder}` does not exist. Creating it now...")
        os.makedirs(outputs_folder)
        logger.info(f"✅ Folder `{outputs_folder}` has been created.")
    else:
        logger.info(f"✅ Folder `{outputs_folder}` already exists.")


    return (
        input_files_folder,
        xlsx_templates_folder,
        outputs_folder,
        report_date,
        config_path,
    )


def main():
    """Main function for batch processing Excel files."""
    for _ in range(2): logger.info("")
    logger.info("-" * 50)
    logger.info("Running Main")
    logger.info("-" * 50)
    
    try:
        # Extract and validate arguments
        input_files_folder, xlsx_templates_folder, outputs_folder, report_date, config_path = parse_args()

        # Load Configuration
        config = config_loader(config_path)

        # Load input data
        input_data_dict = input_data_loader(input_files_folder, config)

        # Add data to output excel files
        add_data_to_files(config['output_from_input_dict'], input_data_dict, xlsx_templates_folder, outputs_folder, report_date)

    except Exception as e:
        logger.error(f"❌ An error occurred: {e}", exc_info=True)

# 🔹 Run the script
if __name__ == "__main__":
    main()
