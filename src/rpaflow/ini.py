"""Módulo INI para rpaflow - Leitura e escrita de arquivos .ini."""

import configparser
from typing import Optional


class Ini:
    """Classe para leitura e escrita de arquivos .ini.

    Usa configparser (built-in). Nenhuma dependência extra.
    """

    def __init__(self, filepath: str, encoding: str = "utf-8"):
        """Abre ou cria arquivo .ini.

        Args:
            filepath: Caminho do arquivo .ini
            encoding: Encoding do arquivo (padrão: utf-8)
        """
        self._filepath = filepath
        self._encoding = encoding
        self._config = configparser.ConfigParser()
        self._config.optionxform = str  # Preserva case das chaves

        try:
            self._config.read(filepath, encoding=encoding)
        except FileNotFoundError:
            pass

    # ========== SEÇÕES ==========

    def get_sections(self) -> list:
        """Retorna lista de seções."""
        return self._config.sections()

    def has_section(self, section: str) -> bool:
        """Verifica se seção existe."""
        return self._config.has_section(section)

    def add_section(self, section: str) -> bool:
        """Adiciona nova seção."""
        if not self._config.has_section(section):
            self._config.add_section(section)
        return True

    def remove_section(self, section: str) -> bool:
        """Remove seção."""
        return self._config.remove_section(section)

    # ========== ITENS ==========

    def get_items(self, section: str) -> list:
        """Retorna lista de tuplas (chave, valor) da seção."""
        return self._config.items(section)

    def has_item(self, section: str, key: str) -> bool:
        """Verifica se item existe na seção."""
        return self._config.has_option(section, key)

    # ========== VALORES ==========

    def get(self, section: str, key: str, fallback=None) -> Optional[str]:
        """Retorna valor como string."""
        return self._config.get(section, key, fallback=fallback)

    def get_int(self, section: str, key: str, fallback: int = 0) -> int:
        """Retorna valor como inteiro."""
        return self._config.getint(section, key, fallback=fallback)

    def get_bool(self, section: str, key: str, fallback: bool = False) -> bool:
        """Retorna valor como booleano.

        Aceita: true/false, yes/no, on/off, 1/0
        """
        return self._config.getboolean(section, key, fallback=fallback)

    def get_float(self, section: str, key: str, fallback: float = 0.0) -> float:
        """Retorna valor como float."""
        return self._config.getfloat(section, key, fallback=fallback)

    # ========== ESCREVER ==========

    def set(self, section: str, key: str, value) -> bool:
        """Define valor de um item.

        Args:
            section: Nome da seção
            key: Nome da chave
            value: Valor (será convertido para string)
        """
        if not self._config.has_section(section):
            self._config.add_section(section)
        self._config.set(section, key, str(value))
        return True

    def remove_item(self, section: str, key: str) -> bool:
        """Remove item de uma seção."""
        return self._config.remove_option(section, key)

    def save(self, filepath: str = None) -> bool:
        """Salva arquivo .ini.

        Args:
            filepath: Caminho alternativo (se None, salva no arquivo original)
        """
        path = filepath or self._filepath
        with open(path, "w", encoding=self._encoding) as f:
            self._config.write(f)
        return True

    # ========== UTILITÁRIOS ==========

    def print_all(self) -> None:
        """Imprime todas as seções e itens."""
        for section in self._config.sections():
            print(f"[{section}]")
            for key, value in self._config.items(section):
                print(f"  {key} = {value}")
            print()

    def to_dict(self) -> dict:
        """Retorna conteúdo como dicionário."""
        result = {}
        for section in self._config.sections():
            result[section] = dict(self._config.items(section))
        return result
