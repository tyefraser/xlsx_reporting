import os
from datetime import datetime
from logger_config import logger

def validate_folder(folder_path):
    if os.path.exists(folder_path) and os.path.isdir(folder_path):
        logger.info(f"✅ Folder `{folder_path}` exists!")
    else:
        logger.error(f"❌ Folder `{folder_path}` does NOT exist!")


def validate_file(file_path):
    if os.path.exists(file_path) and os.path.isfile(file_path):
        logger.info(f"✅ File `{file_path}` exists!")
    else:
        logger.error(f"❌ File `{file_path}` does NOT exist!")


def is_valid_date(date_str):
    try:
        datetime.strptime(date_str, "%Y-%m-%d")  # Ensure correct format
        logger.info(f"✅ Date `{date_str}` is in the correct format (yyyy-mm-dd)!")
        return True
    except ValueError:
        logger.error(f"❌ Date `{date_str}` is NOT in the correct format (yyyy-mm-dd)!")
        return False

