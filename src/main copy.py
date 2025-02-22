import os
import logging
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

def load_input_data(input_files_folder, input_data_dict):

    logger.info("-" * 50)
    logger.info("Loading input data required")

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