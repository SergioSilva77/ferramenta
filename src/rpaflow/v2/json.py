"""Módulo JSON v2 para rpaflow - Leitura e escrita de JSON com dot notation.

Diferenças em relação ao `rpaflow.json.Json` (v1), as principais:

- JSON cujo topo é uma LISTA (`[{...}, {...}]`) carrega: o v1 fazia `Box(lista)` e estourava.
- Arquivo com BOM carrega (`utf-8-sig`).
- Acentos saem iguais em `save()` e `dumps()` (no v1 o Box escapava `\\u00e9` e o dict não).
- `datetime`, `Decimal`, `Path`, `set` etc. são serializados (o v1 estourava `TypeError`).
- `save()` é atômico e cria a pasta; um JSON pela metade nunca fica no disco.
- `default_box` é parâmetro: com `default_box=True` (padrão, igual ao v1) chave inexistente devolve um Box
  vazio em vez de `KeyError` — prático, mas esconde erro de digitação. Use `Json(default_box=False)`
  para ser estrito.
"""

from __future__ import annotations

import json as _json
import os
import tempfile
from typing import Optional


def _padrao(obj):
    """Serializador para o que o json padrão não conhece."""
    if hasattr(obj, "isoformat"):
        return obj.isoformat()
    if isinstance(obj, (set, frozenset)):
        return sorted(obj, key=str)
    if hasattr(obj, "to_dict"):
        return obj.to_dict()
    return str(obj)  # Decimal, Path, UUID...


class Json:
    """Classe para leitura e escrita de arquivos JSON.

    Suporta dot notation via python-box (opcional). Sem python-box, devolve dict/list comuns.
    """

    def __init__(self, default_box: bool = True):
        self._default_box = default_box
        try:
            from box import Box, BoxList
            self._Box = Box
            self._BoxList = BoxList
            self._has_box = True
        except ImportError:
            self._has_box = False

    # ========== CARREGAR ==========

    def _envolver(self, data):
        if not self._has_box:
            return data
        if isinstance(data, dict):
            return self._Box(data, default_box=self._default_box)
        if isinstance(data, list):
            return self._BoxList(data, box_class=self._Box, default_box=self._default_box)
        return data  # número/texto/bool no topo

    def load(self, filepath: str, encoding: str = "utf-8-sig"):
        """Carrega arquivo JSON. Retorna Box/BoxList (dot notation) ou dict/list."""
        with open(filepath, "r", encoding=encoding) as f:
            return self._envolver(_json.load(f))

    def loads(self, text: str):
        """Parse string JSON. Retorna Box/BoxList (dot notation) ou dict/list."""
        return self._envolver(_json.loads(text))

    # ========== SALVAR ==========

    def _simples(self, data):
        """Box/BoxList -> dict/list puros."""
        if self._has_box:
            if isinstance(data, self._Box):
                return data.to_dict()
            if isinstance(data, self._BoxList):
                return data.to_list()
        return data

    def dumps(self, data, indent: Optional[int] = 2) -> str:
        """Serializa dados para string JSON (UTF-8 legível, sem escapar acentos)."""
        return _json.dumps(self._simples(data), indent=indent, ensure_ascii=False, default=_padrao)

    def save(self, filepath: str, data, indent: Optional[int] = 2, encoding: str = "utf-8") -> bool:
        """Salva dados em arquivo JSON (atômico: grava num temporário e troca)."""
        texto = self.dumps(data, indent=indent)
        path = os.path.abspath(filepath)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        fd, tmp = tempfile.mkstemp(dir=os.path.dirname(path), suffix=".tmp")
        try:
            with os.fdopen(fd, "w", encoding=encoding) as f:
                f.write(texto)
            os.replace(tmp, path)
        except Exception:
            if os.path.exists(tmp):
                os.remove(tmp)
            raise
        return True

    # ========== UTILITÁRIOS ==========

    def get_keys(self, data) -> list:
        """Retorna lista de chaves do objeto."""
        return list(data.keys())

    def get_values(self, data) -> list:
        """Retorna lista de valores do objeto."""
        return list(data.values())

    def get_items(self, data) -> list:
        """Retorna lista de tuplas (chave, valor) do objeto."""
        return list(data.items())

    def has_key(self, data, key: str) -> bool:
        """Verifica se chave existe no objeto (sem criar a chave, mesmo com default_box)."""
        return key in data

    def get_value(self, data, key: str, default=None):
        """Retorna valor de uma chave com fallback. Aceita caminho com ponto: get_value(d, 'a.b.c')."""
        if isinstance(data, dict) and key in data:  # chave que tem ponto no nome de verdade
            return data[key]
        atual = data
        for parte in str(key).split("."):
            if isinstance(atual, dict) and parte in atual:
                atual = atual[parte]
            elif isinstance(atual, (list, tuple)) and parte.lstrip("-").isdigit() and -len(atual) <= int(parte) < len(atual):
                atual = atual[int(parte)]
            else:
                return default
        return atual

    def to_dict(self, data) -> dict:
        """Converte para dict puro."""
        return self._simples(data) if isinstance(self._simples(data), dict) else dict(data)

    def to_list(self, data) -> list:
        """Converte para list pura."""
        return self._simples(data) if isinstance(self._simples(data), list) else list(data)
