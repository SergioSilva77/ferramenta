"""Módulo Excel (openpyxl) v2 para rpaflow — planilhas .xlsx sem precisar do Excel instalado.

Diferenças em relação ao `rpaflow.excel.Excel` (v1), as principais:

- `write()` funcionava só no papel: o v1 fazia `ws[range].min_row`, e `ws["A1"]` é uma célula (não tem
  `min_row`), então o método sempre estourava. Agora entende "A1" e "A1:C3".
- `read("A1")` devolvia uma célula e quebrava na iteração; agora sempre devolve lista de listas.
- `FileError` agora é importado (no v1 qualquer erro virava `NameError`).
- Arquivo inexistente NÃO cria planilha vazia em silêncio (typo no caminho virava dado perdido):
  use `create_if_missing=True` se for o que você quer.
- Sheet inexistente dá erro claro, `data_only` para ler o valor calculado das fórmulas, `sheet_names`,
  `create_sheet`, e `with Excel(...) as xl:`.
"""

from __future__ import annotations

import os
from typing import Optional

from rpaflow._exceptions import FileError


class Excel:
    """Classe para operações com planilhas Excel (openpyxl)."""

    def __init__(self, filepath: str = "", create_if_missing: bool = False, data_only: bool = False):
        """
        Args:
            create_if_missing: se o arquivo não existir, começa uma planilha nova (salva em `filepath`).
            data_only: lê o ÚLTIMO VALOR CALCULADO das fórmulas em vez do texto da fórmula.
                Atenção: salvar um arquivo aberto assim PERDE as fórmulas.
        """
        self.filepath = filepath
        self.create_if_missing = create_if_missing
        self.data_only = data_only
        self._workbook = None

    def __enter__(self) -> "Excel":
        self.open()
        return self

    def __exit__(self, *exc) -> bool:
        self.close()
        return False

    # ========== ARQUIVO ==========

    def open(self) -> bool:
        """Abre uma planilha Excel."""
        import openpyxl

        if not self.filepath or not os.path.exists(self.filepath):
            if self.create_if_missing:
                self._workbook = openpyxl.Workbook()
                return True
            raise FileError(f"Planilha não encontrada: {self.filepath!r} (use create_if_missing=True para criar).")
        try:
            self._workbook = openpyxl.load_workbook(self.filepath, data_only=self.data_only)
            return True
        except Exception as e:
            raise FileError(f"Erro ao abrir planilha: {e}") from e

    def _wb(self):
        if not self._workbook:
            raise FileError("Planilha não aberta. Chame open() primeiro.")
        return self._workbook

    def _ws(self, sheet: Optional[str]):
        wb = self._wb()
        if sheet is None:
            return wb.active
        if sheet not in wb.sheetnames:
            raise FileError(f"Sheet '{sheet}' não existe. Sheets: {wb.sheetnames}")
        return wb[sheet]

    @property
    def sheet_names(self) -> list:
        """Nomes das sheets."""
        return list(self._wb().sheetnames)

    def create_sheet(self, name: str) -> bool:
        """Cria uma sheet (se já existir, não faz nada)."""
        wb = self._wb()
        if name not in wb.sheetnames:
            wb.create_sheet(name)
        return True

    # ========== LEITURA / ESCRITA ==========

    def read(self, sheet: Optional[str] = None, range: Optional[str] = None) -> list:
        """Lê dados de uma sheet. Sempre devolve lista de listas ("A1" vira [[valor]])."""
        from openpyxl.utils.cell import range_boundaries

        ws = self._ws(sheet)
        if range is None:
            return [list(row) for row in ws.iter_rows(values_only=True)]
        min_col, min_row, max_col, max_row = range_boundaries(range if ":" in range else f"{range}:{range}")
        return [
            list(row)
            for row in ws.iter_rows(
                min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col, values_only=True
            )
        ]

    def write(self, sheet: Optional[str], range: str, values: list) -> bool:
        """Escreve um bloco (lista de listas) a partir da célula inicial de `range`. Ex: write("Dados", "B2", [[1, 2]])."""
        from openpyxl.utils.cell import range_boundaries

        ws = self._ws(sheet)
        # Aceita "A1" e "A1:C3": o ponto de partida é sempre o canto superior esquerdo.
        coord = range if ":" in range else f"{range}:{range}"
        min_col, min_row, _, _ = range_boundaries(coord)

        if values and not isinstance(values[0], (list, tuple)):
            values = [values]  # lista simples = uma linha
        for i, row in enumerate(values):
            for j, value in enumerate(row):
                ws.cell(row=min_row + i, column=min_col + j, value=value)
        return True

    def append(self, sheet: Optional[str], row: list) -> bool:
        """Adiciona uma linha no fim da sheet."""
        self._ws(sheet).append(list(row))
        return True

    def save(self, filepath: Optional[str] = None) -> bool:
        """Salva a planilha. Cria a pasta de destino se precisar."""
        destino = filepath or self.filepath
        if not destino:
            raise FileError("Informe o caminho: save('arquivo.xlsx') ou Excel('arquivo.xlsx').")
        pasta = os.path.dirname(os.path.abspath(destino))
        os.makedirs(pasta, exist_ok=True)
        self._wb().save(destino)
        return True

    def close(self) -> None:
        """Fecha a planilha."""
        if self._workbook:
            self._workbook.close()
            self._workbook = None
