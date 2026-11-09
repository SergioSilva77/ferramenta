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

from rpaflow import sql, excel, browser, files, api, email, ini

try:
    from importlib.metadata import version as _version
    __version__ = _version("rpaflow")
except Exception:  # rodando direto do código-fonte, sem instalar
    __version__ = "0.2.0"

__all__ = [
    "v2",
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
    "log",
    "ini",
    "json",
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

# Log requer loguru
try:
    from rpaflow import log
except ImportError:
    pass

# JSON requer python-box
try:
    from rpaflow import json
except ImportError:
    pass
