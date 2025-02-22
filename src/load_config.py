import os
import yaml
from logger_config import logger

class NoDuplicateLoader(yaml.SafeLoader):
    """
    Custom YAML loader that raises an error if duplicate keys are found.
    - Prevents silent overwriting of values.
    - Raises `yaml.constructor.ConstructorError` if a duplicate key is detected.
    """

    def construct_mapping(self, node, deep=False):
        mapping = {}
        for key_node, value_node in node.value:
            key = self.construct_object(key_node, deep=deep)
            if key in mapping:
                raise yaml.constructor.ConstructorError(
                    None, None, f"❌ Duplicate key detected in YAML: {key}", key_node.start_mark
                )
            mapping[key] = self.construct_object(value_node, deep=deep)
        return mapping

def load_yaml_with_duplicate_check(yaml_path):
    """
    Loads a YAML file and raises an error if duplicate keys exist.

    Parameters:
        yaml_path (str): Path to the YAML file.

    Returns:
        dict: Parsed YAML configuration.

    Raises:
        FileNotFoundError: If the YAML file is missing.
        yaml.constructor.ConstructorError: If duplicate keys are found.
    """
    if not os.path.exists(yaml_path):
        raise FileNotFoundError(f"❌ Config file '{yaml_path}' not found.")

    logger.info(f"📂 Loading configuration from: {yaml_path}")

    try:
        with open(yaml_path, "r", encoding="utf-8") as file:
            return yaml.load(file, Loader=NoDuplicateLoader)

    except yaml.constructor.ConstructorError as e:
        logger.error(f"⚠️ YAML Error: {e}")
        raise  # Re-raise to prevent silent failures
    except yaml.YAMLError as e:
        logger.error(f"❌ YAML Parsing Error: {e}")
        raise

def config_loader(config_path="config/settings.yaml"):
    """
    Loads configuration settings from a YAML file and validates duplicate keys.

    Parameters:
        config_path (str): Path to the YAML configuration file.

    Returns:
        dict: Loaded configuration settings.
    """
    try:
        config = load_yaml_with_duplicate_check(config_path)
        logger.info("✅ Configuration loaded successfully.")
        return config
    except Exception as e:
        logger.error(f"❌ Failed to load configuration: {e}")
        raise
