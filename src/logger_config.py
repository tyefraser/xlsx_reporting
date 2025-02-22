import logging
import logging.config
import yaml
from pathlib import Path
from datetime import datetime
from logging.handlers import RotatingFileHandler

# Define log directory relative to the project root
LOG_DIR = Path(__file__).parent.parent / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)  # Ensure logs directory exists

# Generate log filename with timestamp (e.g., logs/2025-02-16_12-30-00.log)
log_filename = datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + ".log"
LOG_FILE_PATH = LOG_DIR / log_filename

def setup_logging(config_path="logs/logging_config.yaml"):
    """
    Sets up logging using a YAML configuration file.
    If YAML loading fails, falls back to basic logging.
    
    Args:
        config_path (str): Path to the logging configuration file.
    """
    try:
        with open(config_path, "r") as file:
            config = yaml.safe_load(file)
        
        # Modify file handler dynamically to use a timestamped log file
        if "handlers" in config and "file_handler" in config["handlers"]:
            config["handlers"]["file_handler"]["filename"] = str(LOG_FILE_PATH)
        
        logging.config.dictConfig(config)
        logging.info(f"✅ Logger initialized. Writing logs to: {LOG_FILE_PATH}")

    except Exception as e:
        print(f"❌ Failed to load logging configuration: {e}")
        logging.basicConfig(level=logging.INFO)  # Fallback basic logging

# Call setup_logging to initialize the logger when imported
setup_logging()

# Get a named logger instance
logger = logging.getLogger("hybrid_logger")

# INFO message to confirm the logger has started
logger.info("Logger successfully configured.")
