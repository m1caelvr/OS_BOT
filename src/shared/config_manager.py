import os
import json
from datetime import datetime, timedelta

CONFIG_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "..", "data", "config.json"
)


def get_default_config():
    now = datetime.now()
    return {
        "INITIAL_DATE": (now - timedelta(hours=1)).strftime("%d/%m/%Y %H:%M"),
        "FINAL_DATE": now.strftime("%d/%m/%Y %H:%M"),
    }


def load_config():
    config_dir = os.path.dirname(CONFIG_FILE)
    if not os.path.exists(config_dir):
        os.makedirs(config_dir)

    if not os.path.exists(CONFIG_FILE):
        default_config = get_default_config()
        save_config(default_config)
        return default_config

    with open(CONFIG_FILE, "r") as file:
        return json.load(file)


def save_config(config, CONSTANTS):
    with open(CONFIG_FILE, "w") as file:
        json.dump(config, file, indent=4)

    CONSTANTS.INITIAL_DATE = config["INITIAL_DATE"]
    CONSTANTS.FINAL_DATE = config["FINAL_DATE"]


def get_config_value(key):
    config = load_config()
    return config.get(key, None)


def set_config_value(key, value):
    config = load_config()
    config[key] = value
    save_config(config)
