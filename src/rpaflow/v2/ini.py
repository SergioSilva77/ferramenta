"""Módulo INI v2 para rpaflow - Leitura e escrita de arquivos .ini.

Diferenças em relação ao `rpaflow.ini.Ini` (v1), as principais:

- Senha/valor com `%` (ex: `senha=abc%123`) NÃO estoura mais `InterpolationSyntaxError`: o v1 usava a
  interpolação padrão do configparser, que trata `%` como sintaxe. Aqui o texto é lido literal.
- Arquivo com BOM (salvo pelo Bloco de Notas ou `Set-Content` do PowerShell) é lido normalmente; no v1 dava
  `MissingSectionHeaderError`. A escrita volta sem BOM.
- `get_int/get_float/get_bool` com valor VAZIO (`porta=`) usam o fallback; o v1 estourava `ValueError`.
- `save()` cria a pasta e grava de forma atômica (não deixa o .ini pela metade se o processo morrer).
- Arquivo inexistente: `Ini(path, must_exist=True)` levanta `FileNotFoundError` em vez de seguir vazio.
- `get_section()` e `get_or_raise()` para config obrigatória.
"""

from __future__ import annotations

import configparser
import os
import tempfile
from typing import Optional


class Ini:
    """Classe para leitura e escrita de arquivos .ini (configparser, sem dependências extras)."""

    def __init__(self, filepath: str, encoding: str = "utf-8", must_exist: bool = False):
        """Abre ou cria arquivo .ini.

        Args:
            filepath: Caminho do arquivo .ini
            encoding: Encoding do arquivo (padrão: utf-8; BOM é tolerado na leitura)
            must_exist: True levanta FileNotFoundError se o arquivo não existir
        """
        self._filepath = filepath
        self._encoding = encoding
        # interpolation=None: `%` é texto comum (senhas!). optionxform=str: preserva o case das chaves.
        self._config = configparser.ConfigParser(interpolation=None)
        self._config.optionxform = str

        if not os.path.exists(filepath):
            if must_exist:
                raise FileNotFoundError(f"Arquivo .ini não encontrado: {filepath}")
            return
        leitura = "utf-8-sig" if encoding.lower().replace("_", "-") in ("utf-8", "utf8") else encoding
        with open(filepath, "r", encoding=leitura) as f:
            self._config.read_file(f)

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

    def get_section(self, section: str) -> dict:
        """Retorna a seção como dict. Seção inexistente = KeyError com mensagem clara."""
        if not self._config.has_section(section):
            raise KeyError(f"Seção [{section}] não existe em {self._filepath}")
        return dict(self._config.items(section))

    # ========== ITENS ==========

    def get_items(self, section: str) -> list:
        """Retorna lista de tuplas (chave, valor) da seção."""
        return self._config.items(section)

    def has_item(self, section: str, key: str) -> bool:
        """Verifica se item existe na seção."""
        return self._config.has_option(section, key)

    # ========== VALORES ==========

    def get(self, section: str, key: str, fallback=None) -> Optional[str]:
        """Retorna valor como string (fallback se a seção/chave não existir)."""
        return self._config.get(section, key, fallback=fallback)

    def get_or_raise(self, section: str, key: str) -> str:
        """Valor obrigatório: ausente ou vazio vira KeyError dizendo exatamente o que falta."""
        valor = self._config.get(section, key, fallback=None)
        if valor is None or valor.strip() == "":
            raise KeyError(f"Configuração obrigatória ausente: [{section}] {key} em {self._filepath}")
        return valor

    def _texto(self, section: str, key: str) -> Optional[str]:
        valor = self._config.get(section, key, fallback=None)
        return None if valor is None or valor.strip() == "" else valor.strip()

    def get_int(self, section: str, key: str, fallback: int = 0) -> int:
        """Retorna valor como inteiro (vazio/ausente = fallback)."""
        valor = self._texto(section, key)
        return fallback if valor is None else int(valor)

    def get_bool(self, section: str, key: str, fallback: bool = False) -> bool:
        """Retorna valor como booleano. Aceita: true/false, yes/no, on/off, 1/0 (vazio/ausente = fallback)."""
        valor = self._texto(section, key)
        if valor is None:
            return fallback
        mapa = {"1": True, "yes": True, "true": True, "on": True, "0": False, "no": False, "false": False, "off": False}
        if valor.lower() not in mapa:
            raise ValueError(f"[{section}] {key}={valor!r} não é booleano (use true/false, yes/no, on/off, 1/0)")
        return mapa[valor.lower()]

    def get_float(self, section: str, key: str, fallback: float = 0.0) -> float:
        """Retorna valor como float (vazio/ausente = fallback). Aceita vírgula decimal: '1,5'."""
        valor = self._texto(section, key)
        return fallback if valor is None else float(valor.replace(",", "."))

    # ========== ESCREVER ==========

    def set(self, section: str, key: str, value) -> bool:
        """Define valor de um item (convertido para string; cria a seção se preciso)."""
        if not self._config.has_section(section):
            self._config.add_section(section)
        self._config.set(section, key, str(value))
        return True

    def remove_item(self, section: str, key: str) -> bool:
        """Remove item de uma seção."""
        return self._config.remove_option(section, key)

    def save(self, filepath: str = None) -> bool:
        """Salva arquivo .ini (atômico: grava num temporário e troca)."""
        path = os.path.abspath(filepath or self._filepath)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        fd, tmp = tempfile.mkstemp(dir=os.path.dirname(path), suffix=".tmp")
        try:
            with os.fdopen(fd, "w", encoding=self._encoding) as f:
                self._config.write(f)
            os.replace(tmp, path)
        except Exception:
            if os.path.exists(tmp):
                os.remove(tmp)
            raise
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
        return {s: dict(self._config.items(s)) for s in self._config.sections()}
