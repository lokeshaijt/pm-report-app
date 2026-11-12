import json
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent.parent
CONFIG_PATH = APP_DIR / "config" / "settings.json"
TEMPLATE_PATH = APP_DIR / "templates" / "Tag_Env_Stock_Template.xlsx"


def load_config(path: Path | str = CONFIG_PATH) -> dict:
    with open(path, encoding="utf-8") as f:
        return json.load(f)
