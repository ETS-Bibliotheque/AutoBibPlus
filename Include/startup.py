import configparser
from collections import deque
from pathlib import Path

from .constants import CONFIG_FILE, RATELIMITS
from .create_config import create_config

# Read/create config file (with fixture for RTFD.io)
config = configparser.ConfigParser()
config.optionxform = str
try:
    if not CONFIG_FILE.exists():
        config = create_config()
    else:
        config.read(CONFIG_FILE)
    KEYS = [k.strip() for k in config.get('Authentication', 'APIKey').split(",")]
    DOCS_PATH = [k.strip() for k in config.get('Docs Path', 'Path').split(",")]
except EOFError:
    pass

# Throttling params
_throttling_params = {k: deque(maxlen=v) for k, v in RATELIMITS.items()}

