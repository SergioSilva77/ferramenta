"""rpaflow v2 — as mesmas classes do rpaflow, com os bugs conhecidos corrigidos.

As classes originais (`rpaflow.browser.Browser`, `rpaflow.excel_com.ExcelCom`, ...) continuam intactas.
Para migrar, troque só o import:

    # antes
    from rpaflow.excel_com import ExcelCom
    # depois
    from rpaflow.v2.excel_com import ExcelCom      # ou: from rpaflow.v2 import ExcelCom

Os módulos são carregados sob demanda: importar `rpaflow.v2` não exige playwright, pywin32, loguru etc.
Cada dependência só é exigida quando você usa a classe correspondente.
"""

from importlib import import_module

from rpaflow._exceptions import (  # noqa: F401
    APIError,
    BrowserError,
    ConnectionError,
    EmailError,
    FileError,
    QueryError,
    RPAFlowError,
)

# nome público -> módulo dentro de rpaflow.v2
_CLASSES = {
    "API": "api",
    "Browser": "browser",
    "BrowserTimeout": "browser",
    "Desktop": "desktop",
    "Email": "email",
    "Excel": "excel",
    "ExcelCom": "excel_com",
    "ExcelComError": "excel_com",
    "Files": "files",
    "Ini": "ini",
    "Json": "json",
    "Log": "log",
    "SQL": "sql",
}

_MODULOS = sorted(set(_CLASSES.values()))

__all__ = sorted(_CLASSES) + [
    "APIError", "BrowserError", "ConnectionError", "EmailError", "FileError", "QueryError", "RPAFlowError",
]


def __getattr__(name):
    if name in _CLASSES:
        return getattr(import_module(f"{__name__}.{_CLASSES[name]}"), name)
    if name in _MODULOS:
        return import_module(f"{__name__}.{name}")
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


def __dir__():
    return sorted(set(globals()) | set(_CLASSES) | set(_MODULOS))
