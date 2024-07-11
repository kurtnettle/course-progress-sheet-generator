import logging
import tomllib
from sys import exit as sys_exit

from sheet_generator.config_helper import validate_config

LOG_FILENAME = "app.log"

logging.basicConfig(
    format="%(asctime)s - %(levelname)s - %(funcName)s : %(message)s",
    handlers=[logging.FileHandler(LOG_FILENAME), logging.StreamHandler()],
    level=logging.INFO,
)

LOGGER = logging.getLogger(__name__)


try:
    with open("config.toml", "rb") as f:
        try:
            conf_data = tomllib.load(f)
            validate_config(LOGGER, conf_data)
        except tomllib.TOMLDecodeError as tomlerr:
            LOGGER.error("Invalid Config. %s", tomlerr)
            sys_exit(0)
except FileNotFoundError as e:
    LOGGER.error("File not found. %s", e)
    sys_exit(0)
