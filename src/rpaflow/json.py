"""Módulo JSON para rpaflow - Leitura e escrita de JSON com dot notation."""

import json as _json
from typing import Optional


class Json:
    """Classe para leitura e escrita de arquivos JSON.

    Inspirado no Newtonsoft.Json (C#). Suporta dot notation via python-box.
    Quando python-box não está instalado, retorna dict comum.
    """

    def __init__(self):
        try:
            from box import Box, BoxList
            self._Box = Box
            self._BoxList = BoxList
            self._has_box = True
        except ImportError:
            self._has_box = False

    # ========== CARREGAR ==========

    def load(self, filepath: str, encoding: str = "utf-8"):
        """Carrega arquivo JSON. Retorna Box (com dot notation) ou dict.

        Args:
            filepath: Caminho do arquivo JSON
            encoding: Encoding do arquivo (padrão: utf-8)

        Returns:
            Box com dot notation (se python-box instalado) ou dict
        """
        with open(filepath, "r", encoding=encoding) as f:
            data = _json.load(f)
        if self._has_box:
            return self._Box(data, default_box=True)
        return data

    def loads(self, text: str):
        """Parse string JSON. Retorna Box (com dot notation) ou dict.

        Args:
            text: String JSON

        Returns:
            Box com dot notation (se python-box instalado) ou dict
        """
        data = _json.loads(text)
        if self._has_box:
            return self._Box(data, default_box=True)
        return data

    # ========== SALVAR ==========

    def save(self, filepath: str, data, indent: int = 2, encoding: str = "utf-8") -> bool:
        """Salva dados em arquivo JSON.

        Args:
            filepath: Caminho do arquivo JSON
            data: Dados para salvar (Box, dict, list)
            indent: Indentação (padrão: 2)
            encoding: Encoding do arquivo (padrão: utf-8)
        """
        if self._has_box and isinstance(data, self._Box):
            data.to_json(filename=filepath, indent=indent, encoding=encoding)
        else:
            plain = data.to_dict() if (self._has_box and hasattr(data, 'to_dict')) else data
            with open(filepath, "w", encoding=encoding) as f:
                _json.dump(plain, f, indent=indent, ensure_ascii=False)
        return True

    # ========== SERIALIZAR ==========

    def dumps(self, data, indent: int = 2) -> str:
        """Serializa dados para string JSON.

        Args:
            data: Dados para serializar (Box, dict, list)
            indent: Indentação (padrão: 2)

        Returns:
            String JSON
        """
        if self._has_box and isinstance(data, self._Box):
            return data.to_json(indent=indent)
        plain = data.to_dict() if (self._has_box and hasattr(data, 'to_dict')) else data
        return _json.dumps(plain, indent=indent, ensure_ascii=False)

    # ========== UTILITÁRIOS ==========

    def get_keys(self, data) -> list:
        """Retorna lista de chaves do objeto."""
        if self._has_box and hasattr(data, 'keys'):
            return list(data.keys())
        return list(data.keys())

    def get_values(self, data) -> list:
        """Retorna lista de valores do objeto."""
        if self._has_box and hasattr(data, 'values'):
            return list(data.values())
        return list(data.values())

    def get_items(self, data) -> list:
        """Retorna lista de tuplas (chave, valor) do objeto."""
        if self._has_box and hasattr(data, 'items'):
            return list(data.items())
        return list(data.items())

    def has_key(self, data, key: str) -> bool:
        """Verifica se chave existe no objeto."""
        return key in data

    def get_value(self, data, key: str, default=None):
        """Retorna valor de uma chave com fallback."""
        if hasattr(data, 'get'):
            return data.get(key, default)
        return default

    def to_dict(self, data) -> dict:
        """Converte para dict."""
        if self._has_box and hasattr(data, 'to_dict'):
            return data.to_dict()
        return dict(data)

    def to_list(self, data) -> list:
        """Converte para list."""
        if self._has_box and hasattr(data, 'to_list'):
            return data.to_list()
        return list(data)
