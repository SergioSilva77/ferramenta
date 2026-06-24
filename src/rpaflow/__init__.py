"""rpaflow - Biblioteca Python modular para automação RPA."""

from rpaflow._exceptions import (
    RPAFlowError,
    ConnectionError,
    QueryError,
    FileError,
    BrowserError,
    APIError,
    EmailError,
)

from rpaflow import sql, excel, browser, files, api, email

__version__ = "0.1.0"
__all__ = [
    "RPAFlowError",
    "ConnectionError",
    "QueryError",
    "FileError",
    "BrowserError",
    "APIError",
    "EmailError",
    "sql",
    "excel",
    "excel_com",
    "browser",
    "desktop",
    "files",
    "api",
    "email",
]

# Excel COM só está disponível no Windows
try:
    from rpaflow import excel_com
except ImportError:
    pass

# Desktop requer pyautogui, opencv-python, pillow
try:
    from rpaflow import desktop
except ImportError:
    pass
