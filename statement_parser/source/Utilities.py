import os

from platformdirs import user_data_path


def get_appdata_path(filename: str) -> str:
    """Return a path in the platform's standard per-user app data directory."""
    app_folder = user_data_path("StatementParser", appauthor=False)
    os.makedirs(app_folder, exist_ok=True)
    return os.path.join(app_folder, filename)