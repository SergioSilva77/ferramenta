"""Módulo JSON para rpaflow - Leitura e escrita de JSON com dot notation."""

import json as _json
from typing import Optional


class Json:
    """Classe para leitura e escrita de arquivos JSON.

    Inspirado no Newtonsoft.Json (C#). Suporta dot notation via python-box.
    """

    def __init__(self):
        try:
            from box import Box, BoxList
            self._Box = Box
            self._BoxList = BoxList
        except ImportError:
            raise ImportError(
                "python-box é necessário para o módulo json. "
                "Instale com: pip install rpaflow[java-script]"
            )

    # ========== CARREGAR ==========

    def load(self, filepath: str, encoding: str = "utf-8"):
        """Carrega arquivo JSON com dot notation.

        Args:
            filepath: Caminho do arquivo JSON
            encoding: Encoding do arquivo (padrão: utf-8)

        Returns:
            Box com dot notation
        """
        return self._Box.from_json(filename=filepath, encoding=encoding, default_box=True)

    def loads(self, text: str):
        """Parse string JSON com dot notation.

        Args:
            text: String JSON

        Returns:
            Box com dot notation
        """
        return self._Box.from_json(json_text=text, default_box=True)

    # ========== SALVAR ==========

    def save(self, filepath: str, data, indent: int = 2, encoding: str = "utf-8") -> bool:
        """Salva dados em arquivo JSON.

        Args:
            filepath: Caminho do arquivo JSON
            data: Dados para salvar (Box, dict, list)
            indent: Indentação (padrão: 2)
            encoding: Encoding do arquivo (padrão: utf-8)
        """
        if isinstance(data, self._Box):
            data.to_json(filename=filepath, indent=indent, encoding=encoding)
        else:
            self._Box(data).to_json(filename=filepath, indent=indent, encoding=encoding)
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
        if isinstance(data, self._Box):
            return data.to_json(indent=indent)
        return self._Box(data).to_json(indent=indent)

    # ========== UTILITÁRIOS ==========

    def get_keys(self, data) -> list:
        """Retorna lista de chaves do objeto.

        Args:
            data: Box ou dict

        Returns:
            Lista de chaves
        """
        if isinstance(data, self._Box):
            return list(data.keys())
        return list(data.keys())

    def get_values(self, data) -> list:
        """Retorna lista de valores do objeto.

        Args:
            data: Box ou dict

        Returns:
            Lista de valores
        """
        if isinstance(data, self._Box):
            return list(data.values())
        return list(data.values())

    def get_items(self, data) -> list:
        """Retorna lista de tuplas (chave, valor) do objeto.

        Args:
            data: Box ou dict

        Returns:
            Lista de tuplas (chave, valor)
        """
        if isinstance(data, self._Box):
            return list(data.items())
        return list(data.items())

    def has_key(self, data, key: str) -> bool:
        """Verifica se chave existe no objeto.

        Args:
            data: Box ou dict
            key: Nome da chave

        Returns:
            True se existe, False caso contrário
        """
        if isinstance(data, self._Box):
            return key in data
        return key in data

    def get_value(self, data, key: str, default=None):
        """Retorna valor de uma chave com fallback.

        Args:
            data: Box ou dict
            key: Nome da chave
            default: Valor padrão se não encontrar

        Returns:
            Valor da chave ou default
        """
        if isinstance(data, self._Box):
            return data.get(key, default)
        return data.get(key, default)

    def to_dict(self, data) -> dict:
        """Converte Box para dict.

        Args:
            data: Box ou dict

        Returns:
            dict
        """
        if isinstance(data, self._Box):
            return data.to_dict()
        return dict(data)

    def to_list(self, data) -> list:
        """Converte BoxList para list.

        Args:
            data: BoxList ou list

        Returns:
            list
        """
        if isinstance(data, self._BoxList):
            return data.to_list()
        return list(data)
