import os
import sys

# Helper to get AppData path for user-writable data
def get_appdata_path(filename: str) -> str:
    """Return the full path to a file in the user's AppData (or local app data) directory."""
    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
        if not base:
            base = os.path.abspath(".")
    else:
        base = os.path.expanduser("~/.local/share")
    if not isinstance(base, str):
        base = os.path.abspath(".")
    app_folder = os.path.join(str(base), "statement_parser")
    os.makedirs(app_folder, exist_ok=True)
    return os.path.join(app_folder, filename)