"""Módulo Excel COM v2 para rpaflow (Windows only).

Diferenças em relação ao `rpaflow.excel_com.ExcelCom` (v1), as principais:

- Cada instância sobe o SEU Excel (`DispatchEx`) e só encerra o PID dele — nada de agarrar o Excel dos outros.
- `quit()` / `with` / `__del__` realmente fecham o Excel (no v1 ele sobrava aberto).
- `close()`/métodos sem workbook aberto dão erro claro, não `'NoneType' has no attribute ...`.
- Corrigidos: `clear_filters` (não limpava nada), `wait_calculation` (método COM não existe),
  sheet "very hidden" contada como visível, `find_row` com `.replace(".0", "")` corrompendo valores,
  `find_row` com coluna inexistente casando a linha 0, `read_range_filtered` com UsedRange deslocado,
  `set_sheet_visible(-2)`, filtros com valores numéricos, pivot sem itens, `save_as` com extensão diferente.
"""

from __future__ import annotations

import os
from contextlib import contextmanager
from typing import Optional

from rpaflow._exceptions import RPAFlowError

_XL_SHEET_VISIBLE = -1
_XL_SHEET_HIDDEN = 0
_XL_SHEET_VERY_HIDDEN = 2

# xlFileFormat por extensão (SaveAs sem FileFormat usa o formato ATUAL do arquivo, ignorando a extensão nova)
_FORMATOS = {
    ".xlsx": 51,
    ".xlsm": 52,
    ".xlsb": 50,
    ".xls": 56,
    ".csv": 6,
    ".txt": -4158,
    ".pdf": None,  # tratado à parte (ExportAsFixedFormat)
}


class ExcelComError(RPAFlowError):
    """Erro de automação do Excel via COM."""


def _norm(valor) -> str:
    """Normaliza um valor de célula/critério pra comparação: 12345.0, '12345' e ' 12345 ' viram '12345'."""
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    if hasattr(valor, "strftime"):  # datetime/pywintypes.datetime: compara como dd/mm/aaaa (a data que o usuário vê)
        if (valor.hour, valor.minute, valor.second) == (0, 0, 0):
            return valor.strftime("%d/%m/%Y")
        return valor.strftime("%d/%m/%Y %H:%M:%S")
    texto = str(valor).strip()
    # Só tira ".0" FINAL de número inteiro. O v1 fazia .replace(".0", "") no texto todo: "10.05" virava "105".
    if texto.endswith(".0") and texto[:-2].lstrip("-").isdigit():
        return texto[:-2]
    return texto


_texto_filtro = _norm


def _tupla_2d(raw):
    """O `.Value` do COM devolve escalar, tupla (1 linha/coluna única) ou tupla de tuplas. Padroniza."""
    if raw is None:
        return ()
    if not isinstance(raw, tuple):
        return ((raw,),)
    if raw and not isinstance(raw[0], tuple):
        return (raw,)
    return raw


class ExcelCom:
    """Classe para automação do Excel via COM (win32com).

    Requer Windows com Excel instalado.

    Args:
        visible: mostra a janela do Excel.
        display_alerts: mostra caixas de diálogo (False = responde o padrão sozinho).
        isolated: True (padrão) sobe um Excel próprio. False usa `Dispatch` (pega/compartilha o Excel
            aberto na máquina) — só use se REALMENTE quiser mexer no Excel do usuário; nesse modo
            `quit()` não encerra o processo.

    Uso recomendado:
        with ExcelCom() as xl:
            xl.open("C:/pasta/arquivo.xlsx")
            ...
    """

    def __init__(self, visible: bool = False, display_alerts: bool = False, isolated: bool = True):
        self._isolated = isolated
        self._wb_obj = None
        self._ws_obj = None
        self._xl = None
        if isolated:
            from rpaflow.v2._excel_process import abrir_excel
            self._xl = abrir_excel(visible, display_alerts)
        else:
            import win32com.client
            self._xl = win32com.client.Dispatch("Excel.Application")
            self._xl.Visible = visible
            self._xl.DisplayAlerts = display_alerts

    # ========== ACESSO AO COM ==========

    @property
    def app(self):
        """Objeto COM `Excel.Application`."""
        if self._xl is None:
            raise ExcelComError("Excel já foi encerrado (quit()).")
        return self._xl

    @property
    def workbook(self):
        """Objeto COM do workbook atual."""
        return self._wb

    @property
    def sheet(self):
        """Objeto COM da sheet atual."""
        return self._ws

    @property
    def _wb(self):
        if self._wb_obj is None:
            raise ExcelComError("Nenhum workbook aberto. Chame open() primeiro.")
        return self._wb_obj

    @property
    def _ws(self):
        if self._ws_obj is None:
            raise ExcelComError("Nenhum workbook aberto. Chame open() primeiro.")
        return self._ws_obj

    def __enter__(self) -> "ExcelCom":
        return self

    def __exit__(self, *exc) -> bool:
        self.quit()
        return False

    def __del__(self) -> None:
        # No v1 o Excel sobrava aberto até alguém matar na marra. Aqui cai junto com o objeto.
        try:
            if getattr(self, "_xl", None) is not None:
                self.quit()
        except Exception:
            pass

    # ========== HELPERS INTERNOS ==========

    @contextmanager
    def _janela_visivel(self):
        """Select/Activate precisam da janela visível. Mostra só o tempo necessário e DEVOLVE o estado anterior
        (o v1 deixava o Excel visível pra sempre depois do primeiro select)."""
        was_visible = self._xl.Visible
        if not was_visible:
            self._xl.Visible = True
        try:
            yield
        finally:
            if not was_visible:
                try:
                    self._xl.Visible = False
                except Exception:
                    pass

    def _ensure_full_activation(self):
        """Garante a cadeia completa: App > Workbook > Worksheet (chame dentro de `_janela_visivel`)."""
        self._wb.Activate()
        self._ws.Activate()

    @staticmethod
    def _linhas_visiveis(rng) -> set:
        """Números das linhas (absolutos) NÃO ocultas dentro de `rng`, em poucas chamadas COM.

        `Range.Rows.Hidden`/`EntireRow.Hidden` de um range com linhas mistas devolve False (não "misto"),
        então não serve para detectar filtro. SpecialCells(xlCellTypeVisible) devolve só as áreas visíveis.
        """
        try:
            visiveis = rng.EntireRow.SpecialCells(12)  # xlCellTypeVisible
        except Exception:
            return set()  # nenhuma linha visível
        areas = visiveis.Areas
        linhas = set()
        for i in range(1, areas.Count + 1):
            # Address de EntireRow é "$2:$5" (ou "$2:$2"): 1 chamada COM por área em vez de 2 (Row + Rows.Count).
            partes = areas(i).Address.replace("$", "").split(":")
            ini, fim = int(partes[0]), int(partes[-1])
            linhas.update(range(ini, fim + 1))
        return linhas

    @staticmethod
    def _colunas_visiveis(rng) -> set:
        """Idem `_linhas_visiveis`, para colunas."""
        try:
            visiveis = rng.EntireColumn.SpecialCells(12)
        except Exception:
            return set()
        colunas = set()
        for i in range(1, visiveis.Areas.Count + 1):
            area = visiveis.Areas(i)
            colunas.update(range(area.Column, area.Column + area.Columns.Count))
        return colunas

    def _sheet(self, sheet=None):
        """Resolve uma sheet por nome/índice/objeto COM. None = sheet atual."""
        if sheet is None:
            return self._ws
        if isinstance(sheet, (str, int)):
            return self._wb.Worksheets(sheet)
        return sheet  # já é objeto COM

    def _get_column_letter(self, col: int) -> str:
        """Converte número da coluna para letra. Ex: 1 -> 'A', 27 -> 'AA'"""
        letter = ""
        while col > 0:
            col, remainder = divmod(col - 1, 26)
            letter = chr(65 + remainder) + letter
        return letter

    def _get_column_index(self, table_name: str, column, ignore_case: bool = True) -> int:
        """Converte nome da coluna para índice (1-based). Aceita int ou str.

        Args:
            table_name: Nome da tabela
            column: Índice (int) ou nome (str) da coluna
            ignore_case: Se True, ignora case e espaços (padrão: True)
        """
        if isinstance(column, int):
            return column

        tbl = self._ws.ListObjects(table_name)
        col_name = column.strip() if ignore_case else column

        for i in range(1, tbl.ListColumns.Count + 1):
            tbl_col = tbl.ListColumns(i).Name
            if ignore_case:
                if tbl_col.strip().lower() == col_name.lower():
                    return i
            else:
                if tbl_col == col_name:
                    return i

        raise ValueError(f"Coluna '{column}' não encontrada na tabela '{table_name}'")

    def _get_column_unique_values(self, table_name: str, column, ignore_case: bool = True) -> list:
        """Retorna valores únicos de uma coluna da tabela — bulk read."""
        col_idx = self._get_column_index(table_name, column, ignore_case)
        tbl = self._ws.ListObjects(table_name)
        col_data = tbl.ListColumns(col_idx).DataBodyRange
        if col_data is None:
            return []
        values = {row[0] for row in _tupla_2d(col_data.Value) if row[0] is not None}
        return list(values)

    # ========== WORKBOOK ==========

    def open(self, path: str, readonly: bool = False, password: str = None, update_links: bool = False) -> bool:
        """Abre um workbook (e seleciona a primeira sheet).

        Args:
            readonly: abre somente leitura.
            password: senha de abertura do arquivo.
            update_links: False (padrão) NÃO atualiza vínculos externos — evita o popup "atualizar vínculos?".
        """
        path = os.path.abspath(path)
        if not os.path.exists(path):
            raise FileNotFoundError(f"Arquivo não encontrado: {path}")
        kwargs = {"ReadOnly": readonly, "UpdateLinks": 3 if update_links else 0}
        if password:
            kwargs["Password"] = password
        self._wb_obj = self.app.Workbooks.Open(path, **kwargs)
        self._ws_obj = self._wb_obj.Worksheets(1)
        return True

    def save(self) -> bool:
        """Salva o workbook atual."""
        self._wb.Save()
        return True

    def save_as(self, path: str) -> bool:
        """Salva o workbook em um novo arquivo. O formato vem da extensão (.xlsx, .xlsm, .xlsb, .xls, .csv, .pdf)."""
        path = os.path.abspath(path)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        ext = os.path.splitext(path)[1].lower()
        if ext == ".pdf":
            self._wb.ExportAsFixedFormat(0, path)  # xlTypePDF
        elif ext in _FORMATOS:
            self._wb.SaveAs(path, FileFormat=_FORMATOS[ext])
        else:
            self._wb.SaveAs(path)
        return True

    def close(self, save_changes: bool = True) -> bool:
        """Fecha o workbook atual. Sem workbook aberto não faz nada (não dá erro)."""
        if self._wb_obj is None:
            return False
        try:
            self._wb_obj.Close(SaveChanges=save_changes)
        finally:
            self._wb_obj = None
            self._ws_obj = None
        return True

    def quit(self) -> None:
        """Encerra o Excel. Idempotente. No modo isolado fecha só o processo desta instância."""
        xl, self._xl = self._xl, None
        self._wb_obj = None
        self._ws_obj = None
        if xl is None:
            return
        if self._isolated:
            from rpaflow.v2 import _excel_process
            pid = _excel_process.pedir_quit(xl)
            del xl  # solta a nossa referência COM ANTES de esperar o processo sair
            _excel_process.aguardar_encerramento(pid)
        else:
            try:
                xl.Quit()
            except Exception:
                pass

    # ========== MACRO ==========

    def run_macro(self, name: str, *args):
        """Executa uma macro. Ex: xl.run_macro('NomeDaMacro', arg1, arg2)"""
        return self.app.Application.Run(name, *args)

    # ========== CÉLULA / RANGE ==========

    def get_value(self, cell_range: str):
        """Lê valor de célula ou intervalo. Ex: xl.get_value('A1') ou xl.get_value('A1:C10')"""
        return self._ws.Range(cell_range).Value

    def set_value(self, cell_range: str, value) -> bool:
        """Escreve valor em célula ou intervalo. Aceita lista de listas para escrever um bloco."""
        if isinstance(value, (list, tuple)):
            if value and all(isinstance(r, (list, tuple)) for r in value):
                value = tuple(tuple(r) for r in value)  # bloco: lista de linhas
            else:
                value = (tuple(value),)  # lista simples = UMA linha
        self._ws.Range(cell_range).Value = value
        return True

    def get_cell(self, row: int, col: int):
        """Lê valor por linha/coluna (1-based)."""
        return self._ws.Cells(row, col).Value

    def set_cell(self, row: int, col: int, value) -> bool:
        """Escreve valor por linha/coluna (1-based)."""
        self._ws.Cells(row, col).Value = value
        return True

    def select_cell(self, cell_range: str) -> bool:
        """Seleciona e ativa uma célula ou intervalo."""
        with self._janela_visivel():
            self._ensure_full_activation()
            self._ws.Range(cell_range).Select()
        return True

    def activate_cell(self, cell_range: str) -> bool:
        """Ativa uma célula (sem selecionar o intervalo)."""
        with self._janela_visivel():
            self._ensure_full_activation()
            self._ws.Range(cell_range).Activate()
        return True

    def set_formula(self, cell_range: str, formula: str) -> bool:
        """Insere fórmula. Ex: xl.set_formula('A1', '=SUM(A1:A10)')"""
        self._ws.Range(cell_range).Formula = formula
        return True

    def set_formula_local(self, cell_range: str, formula: str) -> bool:
        """Insere fórmula local (pt-BR). Ex: xl.set_formula_local('A1', '=SOMA(A1:A10)')"""
        self._ws.Range(cell_range).FormulaLocal = formula
        return True

    def clear(self, cell_range: str) -> bool:
        """Limpa conteúdo de células."""
        self._ws.Range(cell_range).ClearContents()
        return True

    # `dest_sheet` aceita nome, índice ou objeto COM (no v1 só objeto COM).
    def copy_range(self, src_range: str, dest_range: str, dest_sheet=None) -> bool:
        """Copia intervalo para outro local."""
        self._ws.Range(src_range).Copy(self._sheet(dest_sheet).Range(dest_range))
        return True

    def _colar_especial(self, src_range: str, dest_range: str, dest_sheet, paste: int) -> bool:
        try:
            self._ws.Range(src_range).Copy()
            self._sheet(dest_sheet).Range(dest_range).PasteSpecial(Paste=paste)
        finally:
            # Sai do modo "formigas marchando" mesmo se a colagem falhar, senão o Excel trava esperando Enter.
            try:
                self._xl.CutCopyMode = False
            except Exception:
                pass
        return True

    def copy_values(self, src_range: str, dest_range: str, dest_sheet=None) -> bool:
        """Copia apenas valores (sem fórmulas)."""
        return self._colar_especial(src_range, dest_range, dest_sheet, -4163)  # xlPasteValues

    def copy_format(self, src_range: str, dest_range: str, dest_sheet=None) -> bool:
        """Copia apenas a formatação."""
        return self._colar_especial(src_range, dest_range, dest_sheet, -4122)  # xlPasteFormats

    def move_range(self, src_range: str, dest_range: str, dest_sheet=None) -> bool:
        """Move (recorta e cola) o range para outro local."""
        self._ws.Range(src_range).Cut(self._sheet(dest_sheet).Range(dest_range))
        return True

    def auto_fit(self, columns: str = "A:Z") -> bool:
        """Ajusta largura das colunas."""
        self._ws.Columns(columns).AutoFit()
        return True

    def get_cell_address(self, row: int, col: int) -> str:
        """Retorna endereço da célula. Ex: (1, 1) -> 'A1'"""
        return self._ws.Cells(row, col).Address.replace("$", "")

    def get_cell_by_address(self, address: str):
        """Lê valor por endereço. Ex: 'A1'"""
        return self._ws.Range(address).Value

    # ========== SHEET ==========

    def select_sheet(self, name_or_index) -> bool:
        """Seleciona uma sheet por nome ou índice (1-based)."""
        self._ws_obj = self._wb.Worksheets(name_or_index)
        with self._janela_visivel():
            self._ensure_full_activation()
        return True

    def list_sheets(self) -> list:
        """Lista todas as sheets do workbook."""
        return [self._wb.Worksheets(i).Name for i in range(1, self._wb.Worksheets.Count + 1)]

    def add_sheet(self, name: str = None, at_end: bool = False) -> bool:
        """Adiciona uma nova sheet (antes da atual, ou no fim com at_end=True)."""
        if at_end:
            ultima = self._wb.Worksheets(self._wb.Worksheets.Count)
            new_sheet = self._wb.Worksheets.Add(After=ultima)
        else:
            new_sheet = self._wb.Worksheets.Add()
        if name:
            new_sheet.Name = name
        return True

    def rename_sheet(self, old_name: str, new_name: str) -> bool:
        """Renomeia uma sheet."""
        self._wb.Worksheets(old_name).Name = new_name
        return True

    def copy_sheet(self, name: str, before: str = None) -> bool:
        """Copia uma sheet."""
        ws = self._wb.Worksheets(name)
        if before:
            ws.Copy(Before=self._wb.Worksheets(before))
        else:
            ws.Copy()
        return True

    def delete_sheet(self, name: str) -> bool:
        """Deleta uma sheet. Se for a sheet atual, volta pra primeira."""
        ws = self._wb.Worksheets(name)
        era_atual = ws.Name == self._ws.Name
        ws.Delete()
        if era_atual:
            self._ws_obj = self._wb.Worksheets(1)
        return True

    def set_sheet_visible(self, name: str, visible=True) -> bool:
        """Controla visibilidade da sheet.

        Args:
            visible: True (visível), False (oculta) ou "very_hidden" (só reexibe via VBA).
                O v1 mandava -2 pra "very hidden", que não é um valor válido (o certo é 2).
        """
        if visible in ("very_hidden", "veryhidden", _XL_SHEET_VERY_HIDDEN):
            estado = _XL_SHEET_VERY_HIDDEN
        else:
            estado = _XL_SHEET_VISIBLE if visible else _XL_SHEET_HIDDEN
        self._wb.Worksheets(name).Visible = estado
        return True

    def _sheets_por_visibilidade(self, visiveis: bool) -> list:
        # Visible vale -1 (visível), 0 (oculta) ou 2 (very hidden). `not Visible` do v1
        # tratava very hidden (2) como visível.
        nomes = []
        for i in range(1, self._wb.Worksheets.Count + 1):
            ws = self._wb.Worksheets(i)
            if (ws.Visible == _XL_SHEET_VISIBLE) == visiveis:
                nomes.append(ws.Name)
        return nomes

    def count_hidden_sheets(self) -> int:
        """Conta sheets ocultas (inclui very hidden)."""
        return len(self._sheets_por_visibilidade(False))

    def list_hidden_sheets(self) -> list:
        """Lista sheets ocultas (inclui very hidden)."""
        return self._sheets_por_visibilidade(False)

    def list_visible_sheets(self) -> list:
        """Lista sheets visíveis."""
        return self._sheets_por_visibilidade(True)

    # ========== OCULTAR COLUNAS/LINHAS ==========

    def hide_columns(self, columns: str) -> bool:
        """Oculta colunas. Ex: xl.hide_columns('B:D')"""
        self._ws.Columns(columns).Hidden = True
        return True

    def show_columns(self, columns: str) -> bool:
        """Mostra colunas ocultas."""
        self._ws.Columns(columns).Hidden = False
        return True

    def hide_rows(self, start_row: int, end_row: int = None) -> bool:
        """Oculta linhas. Ex: xl.hide_rows(5, 10) ou xl.hide_rows(5)"""
        if end_row:
            self._ws.Rows(f"{start_row}:{end_row}").Hidden = True
        else:
            self._ws.Rows(start_row).Hidden = True
        return True

    def show_rows(self, start_row: int, end_row: int = None) -> bool:
        """Mostra linhas ocultas."""
        if end_row:
            self._ws.Rows(f"{start_row}:{end_row}").Hidden = False
        else:
            self._ws.Rows(start_row).Hidden = False
        return True

    def list_hidden_columns(self) -> list:
        """Lista colunas ocultas (letras) dentro da área usada."""
        ur = self._ws.UsedRange
        primeira, ultima = ur.Column, ur.Column + ur.Columns.Count - 1
        visiveis = self._colunas_visiveis(self._ws.Range(self._ws.Cells(1, primeira), self._ws.Cells(1, ultima)))
        return [self._get_column_letter(c) for c in range(primeira, ultima + 1) if c not in visiveis]

    def list_hidden_rows(self) -> list:
        """Lista linhas ocultas (números) dentro da área usada."""
        ur = self._ws.UsedRange
        primeira, ultima = ur.Row, ur.Row + ur.Rows.Count - 1
        visiveis = self._linhas_visiveis(self._ws.Range(self._ws.Cells(primeira, 1), self._ws.Cells(ultima, 1)))
        return [r for r in range(primeira, ultima + 1) if r not in visiveis]

    # ========== LISTAR TABELAS ==========

    def list_tables(self) -> list:
        """Lista todas as tabelas (ListObjects) da sheet atual."""
        return [self._ws.ListObjects(i).Name for i in range(1, self._ws.ListObjects.Count + 1)]

    # ========== TABELA (ListObject) ==========

    def read_table(self, table_name: str) -> list:
        """Lê corpo da tabela (ListObject) — bulk read."""
        tbl = self._ws.ListObjects(table_name)
        data = tbl.DataBodyRange
        if data is None:
            return []
        return [list(row) for row in _tupla_2d(data.Value)]

    def read_filtered_table(self, table_name: str) -> list:
        """Lê apenas linhas visíveis (não filtradas) da tabela — hybrid read."""
        tbl = self._ws.ListObjects(table_name)
        data = tbl.DataBodyRange
        if data is None:
            return []

        # BULK: ler todos os dados de uma vez
        raw = _tupla_2d(data.Value)

        # Linhas visíveis em poucas chamadas COM (o v1 fazia 1 chamada por linha, e é o gargalo em tabela grande).
        visiveis = self._linhas_visiveis(data)
        primeira = data.Row
        linhas = []
        for r, row in enumerate(raw):
            if primeira + r in visiveis:
                linhas.append([v if v is not None else "" for v in row])
        return linhas

    def count_filtered_rows(self, table_name: str) -> int:
        """Conta linhas visíveis (não filtradas) da tabela."""
        return len(self.read_filtered_table(table_name))

    def read_table_header(self, table_name: str) -> list:
        """Lê cabeçalho da tabela — bulk read."""
        tbl = self._ws.ListObjects(table_name)
        linhas = _tupla_2d(tbl.HeaderRowRange.Value)
        return list(linhas[0]) if linhas else []

    def read_table_column(self, table_name: str, column_name) -> list:
        """Lê uma coluna da tabela pelo nome (ou índice 1-based) — bulk read. Tabela vazia devolve []."""
        col_idx = self._get_column_index(table_name, column_name)
        tbl = self._ws.ListObjects(table_name)
        col = tbl.ListColumns(col_idx).DataBodyRange
        if col is None:
            return []
        return [row[0] for row in _tupla_2d(col.Value)]

    def count_table_rows(self, table_name: str) -> int:
        """Conta linhas da tabela (0 se estiver vazia — o v1 dava erro)."""
        tbl = self._ws.ListObjects(table_name)
        body = tbl.DataBodyRange
        return 0 if body is None else body.Rows.Count

    def get_table_header_row(self, table_name: str) -> int:
        """Retorna a linha (1-based) onde começa o header da tabela."""
        tbl = self._ws.ListObjects(table_name)
        return tbl.HeaderRowRange.Row

    def read_table_as_dicts(self, table_name: str) -> list:
        """Lê tabela como lista de dicionários."""
        header = self.read_table_header(table_name)
        nomes = [str(h or "").strip() for h in header]
        raw = self.read_table(table_name)
        return [dict(zip(nomes, row)) for row in raw]

    def read_filtered_table_as_dicts(self, table_name: str) -> list:
        """Lê tabela filtrada como lista de dicionários."""
        header = self.read_table_header(table_name)
        nomes = [str(h or "").strip() for h in header]
        raw = self.read_filtered_table(table_name)
        return [dict(zip(nomes, row)) for row in raw]

    def find_row(self, table_name: str, criteria: dict, data: list = None) -> int:
        """Retorna índice 0-based da primeira linha que bate com criteria.

        Args:
            table_name: Nome da tabela
            criteria: Dicionário com coluna=valor. Ex: {"NF": "12345", "Identificador": "678"}
            data: Lista de listas já lida (opcional). Se não informado, lê a tabela.

        Returns:
            Índice 0-based ou None se não encontrar.
        """
        header = self.read_table_header(table_name)
        nomes = [str(h or "").strip() for h in header]
        if data is None:
            data = self.read_table(table_name)

        # O v1 ignorava em silêncio critério com coluna inexistente — e se NENHUM existisse, devolvia a linha 0.
        indices = {}
        for k in criteria:
            if k not in nomes:
                raise KeyError(f"Coluna '{k}' não existe na tabela '{table_name}'. Colunas: {nomes}")
            indices[k] = nomes.index(k)
        if not indices:
            raise ValueError("criteria vazio: informe ao menos uma coluna=valor")

        alvo = {k: _norm(v) for k, v in criteria.items()}
        for idx, row in enumerate(data):
            if all(_norm(row[indices[k]] if indices[k] < len(row) else None) == alvo[k] for k in criteria):
                return idx
        return None

    def set_table_cell(self, table_name: str, data_row: int, column, value) -> bool:
        """Escreve numa célula do corpo da tabela. data_row é 0-based."""
        header_row = self.get_table_header_row(table_name)
        col_idx = self._get_column_index(table_name, column) if isinstance(column, str) else column
        self.set_cell(header_row + 1 + data_row, col_idx, value)
        return True

    def get_table_cell(self, table_name: str, data_row: int, column):
        """Lê uma célula do corpo da tabela. data_row é 0-based."""
        header_row = self.get_table_header_row(table_name)
        col_idx = self._get_column_index(table_name, column) if isinstance(column, str) else column
        return self.get_cell(header_row + 1 + data_row, col_idx)

    def read_range_filtered(self, header_row: int, start_row: int = None,
                            filters: dict = None, sheet: str = None) -> list:
        """Lê um range e filtra em Python — bulk read (para planilhas sem ListObject).

        Args:
            header_row: Linha do cabeçalho (1-based)
            start_row: Linha inicial dos dados (padrão: header_row + 1)
            filters: Dicionário com coluna=valor. Ex: {"Status": "A Pagar", "Data": "22/06/2026"}
            sheet: Nome da sheet (opcional)

        Returns:
            Lista de dicionários com as linhas que atendem ao filtro.
        """
        ws = self._sheet(sheet)
        if start_row is None:
            start_row = header_row + 1

        # UsedRange pode NÃO começar na linha/coluna 1: "Rows.Count" sozinho (v1) não é a última linha.
        ur = ws.UsedRange
        last_col = ur.Column + ur.Columns.Count - 1
        last_row = ur.Row + ur.Rows.Count - 1

        if start_row > last_row:
            return []

        # BULK: ler cabeçalho de uma vez
        header_raw = _tupla_2d(ws.Range(ws.Cells(header_row, 1), ws.Cells(header_row, last_col)).Value)
        header = [("" if h is None else str(h)).strip() for h in (header_raw[0] if header_raw else ())]

        # Coluna de filtro que não existe = erro (o v1 ignorava e devolvia tudo).
        filter_cols = {}
        for col_name in (filters or {}):
            if col_name not in header:
                raise KeyError(f"Coluna '{col_name}' não existe no cabeçalho (linha {header_row}): {header}")
            filter_cols[col_name] = header.index(col_name)

        # BULK: ler todos os dados de uma vez
        data = _tupla_2d(ws.Range(ws.Cells(start_row, 1), ws.Cells(last_row, last_col)).Value)
        if not data:
            return []

        # Linhas ocultas (autofiltro/ocultar manual). `Rows.Hidden` num range misto devolve False — o v1 lia
        # isso como "nenhuma oculta" e trazia linhas filtradas. Aqui vem das áreas visíveis.
        visiveis = self._linhas_visiveis(ws.Range(ws.Cells(start_row, 1), ws.Cells(last_row, 1)))
        hidden_flags = [(start_row + i) not in visiveis for i in range(len(data))]

        filtros = {k: (None if v is None else _norm(v)) for k, v in (filters or {}).items()}

        # Filtrar em Python (zero COM)
        result = []
        for i, row in enumerate(data):
            if hidden_flags[i]:
                continue

            row_dict = {nome: (row[c] if c < len(row) else None) for c, nome in enumerate(header)}

            # filtros: coluna=valor. valor None significa "coluna preenchida".
            ok = True
            for col_name, valor in filtros.items():
                cell_val = _norm(row_dict.get(col_name))
                if valor is None:
                    if cell_val == "":
                        ok = False
                elif cell_val != valor:
                    ok = False
                if not ok:
                    break
            if ok:
                result.append(row_dict)

        return result

    def refresh_table(self, table_name: str) -> bool:
        """Atualiza uma tabela ligada a uma consulta externa (QueryTable)."""
        tbl = self._ws.ListObjects(table_name)
        try:
            qt = tbl.QueryTable
        except Exception:
            raise ExcelComError(
                f"A tabela '{table_name}' não tem consulta externa (QueryTable); não há o que atualizar. "
                "Use refresh_all() se ela depende de outra fonte."
            )
        qt.Refresh(BackgroundQuery=False)
        return True

    def refresh_all(self, wait: bool = True) -> bool:
        """Atualiza todas as conexões, tabelas e dinâmicas. wait=True espera consultas assíncronas terminarem."""
        self._wb.RefreshAll()
        if wait:
            self.wait_calculation()
        return True

    # ========== FILTROS DE TABELA ==========

    def filter_column(self, table_name: str, column, criteria: str, ignore_case: bool = True) -> bool:
        """Filtra coluna. Ex: xl.filter_column('Vendas', 'Status', 'Aprovado')"""
        col_idx = self._get_column_index(table_name, column, ignore_case)
        tbl = self._ws.ListObjects(table_name)
        tbl.Range.AutoFilter(Field=col_idx, Criteria1=criteria)
        return True

    def filter_column_values(self, table_name: str, column, values: list, ignore_case: bool = True) -> bool:
        """Filtra por lista de valores. Ex: xl.filter_column_values('Vendas', 'Tipo', ['PCD', 'PJ'])"""
        col_idx = self._get_column_index(table_name, column, ignore_case)
        tbl = self._ws.ListObjects(table_name)
        # O AutoFilter por lista só aceita TEXTO: com número/data o v1 estourava. Aqui tudo vira o texto exibido.
        criterios = [_texto_filtro(v) for v in values]
        if not criterios:
            raise ValueError("values vazio: nada ficaria visível. Use filter_column_blanks/clear_filters.")
        tbl.Range.AutoFilter(Field=col_idx, Criteria1=criterios, Operator=7)  # xlFilterValues
        return True

    def filter_column_exclude(self, table_name: str, column, values: list, ignore_case: bool = True) -> bool:
        """Esconde valores específicos."""
        col_idx = self._get_column_index(table_name, column, ignore_case)
        all_values = self._get_column_unique_values(table_name, col_idx)
        excluir = {_norm(v) for v in values}
        visible = [v for v in all_values if _norm(v) not in excluir]
        if not visible:
            raise ValueError("Todos os valores da coluna seriam escondidos (o Excel não permite filtro vazio).")
        return self.filter_column_values(table_name, col_idx, visible)

    def filter_column_number(self, table_name: str, column, criteria: str, ignore_case: bool = True) -> bool:
        """Filtro numérico. Ex: xl.filter_column_number('Vendas', 'Valor', '>1000')"""
        col_idx = self._get_column_index(table_name, column, ignore_case)
        tbl = self._ws.ListObjects(table_name)
        tbl.Range.AutoFilter(Field=col_idx, Criteria1=criteria)
        return True

    def filter_column_color(self, table_name: str, column, color: int, type: str = "fill", ignore_case: bool = True) -> bool:
        """Filtra por cor. Ex: xl.filter_column_color('Vendas', 'Status', 0xFFFF, 'fill')"""
        col_idx = self._get_column_index(table_name, column, ignore_case)
        tbl = self._ws.ListObjects(table_name)
        operator = 8 if type == "fill" else 10  # xlFilterCellColor=8, xlFilterFontColor=10
        tbl.Range.AutoFilter(Field=col_idx, Operator=operator, Criteria1=color)
        return True

    def filter_column_blanks(self, table_name: str, column, exclude_empty: bool = True, ignore_case: bool = True) -> bool:
        """Filtra vazios. exclude_empty=True mostra só não-vazios."""
        col_idx = self._get_column_index(table_name, column, ignore_case)
        tbl = self._ws.ListObjects(table_name)
        if exclude_empty:
            tbl.Range.AutoFilter(Field=col_idx, Criteria1="<>")
        else:
            tbl.Range.AutoFilter(Field=col_idx, Criteria1="=")
        return True

    def clear_filters(self, table_name: str) -> bool:
        """Limpa todos os filtros da tabela."""
        # O v1 fazia `tbl.Range.ShowAllData`, que não existe em Range: dava AttributeError, o `except: pass`
        # engolia e NENHUM filtro era limpo. ShowAllData é método da Worksheet (e do AutoFilter da tabela).
        tbl = self._ws.ListObjects(table_name)
        try:
            af = tbl.AutoFilter
        except Exception:
            af = None
        if af is not None:
            try:
                if af.FilterMode:
                    af.ShowAllData()
            except Exception:
                pass
        else:
            try:
                if self._ws.FilterMode:
                    self._ws.ShowAllData()
            except Exception:
                pass
        return True

    def remove_column_filter(self, table_name: str, column, ignore_case: bool = True) -> bool:
        """Remove filtro de uma coluna específica (mostra todos os valores)."""
        col_idx = self._get_column_index(table_name, column, ignore_case)
        tbl = self._ws.ListObjects(table_name)
        tbl.Range.AutoFilter(Field=col_idx)
        return True

    def sort_column(self, table_name: str, column, order: str = "asc", ignore_case: bool = True) -> bool:
        """Classifica coluna. order='asc' ou 'desc'."""
        col_idx = self._get_column_index(table_name, column, ignore_case)
        tbl = self._ws.ListObjects(table_name)
        sort = self._ws.Sort
        sort.SortFields.Clear()
        sort.SortFields.Add(
            Key=tbl.ListColumns(col_idx).Range,
            SortOn=0,
            Order=1 if order == "asc" else 2,
            DataOption=0
        )
        sort.SetRange(tbl.Range)
        sort.Header = 1  # xlYes
        sort.Apply()
        return True

    # ========== TABELA DINÂMICA (PivotTable) ==========

    def list_pivot_tables(self, sheet: str = None) -> list:
        """Lista PivotTables da sheet."""
        ws = self._sheet(sheet)
        return [ws.PivotTables(i).Name for i in range(1, ws.PivotTables().Count + 1)]

    def refresh_pivot_table(self, name: str, sheet: str = None) -> bool:
        """Atualiza PivotTable."""
        ws = self._sheet(sheet)
        ws.PivotTables(name).RefreshTable()
        return True

    def refresh_pivot_caches(self) -> bool:
        """Atualiza o CACHE de todas as dinâmicas do workbook (RefreshTable só redesenha a partir do cache)."""
        caches = self._wb.PivotCaches()
        for i in range(1, caches.Count + 1):
            caches.Item(i).Refresh()
        return True

    def set_pivot_source(self, pivot_name: str, source: str, sheet: str = None, refresh: bool = True) -> bool:
        """Troca a fonte da dinâmica. Prefira o NOME de uma tabela (ListObject) a um endereço fixo.

        Com endereço fixo ("A4:L104") linhas novas ficam de fora e a dinâmica "não atualiza".
        Com o nome da tabela a fonte cresce junto com os dados.
        """
        ws = self._sheet(sheet)
        cache = ws.PivotTables(pivot_name).PivotCache()
        cache.SourceData = source
        if refresh:
            cache.Refresh()
        return True

    def filter_pivot_values(self, pivot_name: str, field_name: str, visible_items: list, sheet: str = None) -> bool:
        """Filtra itens visíveis da PivotTable."""
        ws = self._sheet(sheet)
        pt = ws.PivotTables(pivot_name)
        pf = pt.PivotFields(field_name)
        pf.ClearAllFilters()
        itens = [pf.PivotItems(i) for i in range(1, pf.PivotItems().Count + 1)]
        queridos = {_norm(v) for v in visible_items}
        # Excel não deixa a dinâmica ficar sem NENHUM item visível: valida antes de esconder (o v1 estourava no meio).
        if not any(_norm(pi.Name) in queridos for pi in itens):
            raise ValueError(f"Nenhum dos itens {list(visible_items)} existe no campo '{field_name}'.")
        for pi in itens:
            if _norm(pi.Name) in queridos:
                pi.Visible = True
        for pi in itens:
            if _norm(pi.Name) not in queridos:
                pi.Visible = False
        return True

    def filter_pivot_olap(self, pivot_name: str, field_name: str, value: str, sheet: str = None) -> bool:
        """Filtra PivotTable OLAP usando CurrentPage.
        
        Args:
            pivot_name: Nome da PivotTable
            field_name: Nome do campo OLAP. Ex: "[Base_atualizada].[TipoLead].[TipoLead]"
            value: Valor para filtrar. Ex: "[Base_atualizada].[TipoLead].&[CONSÓRCIO]"
            sheet: Nome da sheet (opcional)
        """
        ws = self._sheet(sheet)
        pt = ws.PivotTables(pivot_name)
        pf = pt.PivotFields(field_name)
        pf.ClearAllFilters()
        pf.CurrentPage = value
        return True

    def filter_pivot_exclude(self, pivot_name: str, field_name: str, exclude_items: list, sheet: str = None) -> bool:
        """Esconde itens específicos da PivotTable."""
        ws = self._sheet(sheet)
        pt = ws.PivotTables(pivot_name)
        pf = pt.PivotFields(field_name)
        pf.ClearAllFilters()
        itens = [pf.PivotItems(i) for i in range(1, pf.PivotItems().Count + 1)]
        excluir = {_norm(v) for v in exclude_items}
        if all(_norm(pi.Name) in excluir for pi in itens):
            raise ValueError(f"Todos os itens do campo '{field_name}' seriam escondidos.")
        for pi in itens:
            if _norm(pi.Name) in excluir:
                pi.Visible = False
        return True

    def clear_pivot_filters(self, pivot_name: str, sheet: str = None) -> bool:
        """Limpa filtros da PivotTable."""
        ws = self._sheet(sheet)
        pt = ws.PivotTables(pivot_name)
        for i in range(1, pt.PivotFields().Count + 1):
            pt.PivotFields(i).ClearAllFilters()
        return True

    def set_pivot_page_filter(self, pivot_name: str, field_name: str, value: str, sheet: str = None) -> bool:
        """Define filtro de página da PivotTable."""
        ws = self._sheet(sheet)
        ws.PivotTables(pivot_name).PivotFields(field_name).CurrentPage = value
        return True

    def get_pivot_field_items(self, pivot_name: str, field_name: str, sheet: str = None) -> list:
        """Lista itens de um campo da PivotTable."""
        ws = self._sheet(sheet)
        pf = ws.PivotTables(pivot_name).PivotFields(field_name)
        return [{"name": pf.PivotItems(i).Name, "visible": pf.PivotItems(i).Visible}
                for i in range(1, pf.PivotItems().Count + 1)]

    # ========== CONEXÕES ==========

    def list_connections(self) -> list:
        """Lista todas as conexões do workbook."""
        return [self._wb.Connections(i).Name for i in range(1, self._wb.Connections.Count + 1)]

    def refresh_connection(self, name: str) -> bool:
        """Atualiza uma conexão pelo nome."""
        self._wb.Connections(name).Refresh()
        return True

    def get_connection_string(self, table_name: str) -> str:
        """Lê a connection string de uma tabela."""
        tbl = self._ws.ListObjects(table_name)
        return tbl.QueryTable.Connection

    def get_connection_sql(self, table_name: str) -> str:
        """Lê o SQL de uma tabela."""
        tbl = self._ws.ListObjects(table_name)
        return tbl.QueryTable.CommandText

    # ========== CÁLCULO ==========

    def calculate(self) -> bool:
        """Calcula todas as fórmulas da planilha."""
        self.app.Calculate()
        return True

    def wait_calculation(self) -> bool:
        """Aguarda queries assíncronas finalizarem.

        O v1 chamava `CalculateUntilAsyncQueriesStop`, método que não existe (o certo é `...Done`).
        """
        self.app.CalculateUntilAsyncQueriesDone()
        return True

    # ========== INFO ==========

    def get_filename(self) -> str:
        """Retorna o nome do arquivo."""
        return self._wb.Name

    def get_filepath(self) -> str:
        """Retorna o caminho completo do arquivo."""
        return self._wb.FullName

    def get_last_row(self, sheet: str = None) -> int:
        """Retorna a última linha usada (número da linha, não a quantidade de linhas do UsedRange)."""
        ws = self._sheet(sheet)
        return ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

    def get_last_col(self, sheet: str = None) -> int:
        """Retorna a última coluna usada (número da coluna)."""
        ws = self._sheet(sheet)
        return ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1

    def get_last_used_row(self, sheet: str = None) -> int:
        """Retorna a última linha com dados (precisa)."""
        ws = self._sheet(sheet)
        return ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

    def get_last_used_row_in_column(self, column: str, sheet: str = None) -> int:
        """Retorna a última linha populada de uma coluna específica."""
        ws = self._sheet(sheet)
        return ws.Cells(ws.Rows.Count, column).End(-4162).Row  # xlUp = -4162

    # ========== APP / WINDOW ==========

    def set_visible(self, visible: bool = True) -> bool:
        """Controla visibilidade do Excel."""
        self.app.Visible = visible
        return True

    def set_display_alerts(self, enabled: bool = True) -> bool:
        """Controla exibição de alertas."""
        self.app.DisplayAlerts = enabled
        return True

    def set_screen_updating(self, enabled: bool = True) -> bool:
        """Controla atualização de tela (performance)."""
        self.app.ScreenUpdating = enabled
        return True

    def set_zoom(self, percent: int = 100) -> bool:
        """Define zoom da janela."""
        self._wb.Windows(1).Zoom = percent
        return True

    def freeze_panes(self, cell: str = None, unfreeze: bool = False) -> bool:
        """Congela painéis. Sem `cell`, congela na seleção atual; com `cell` ('B2'), acima/à esquerda dela.

        O v1 congelava na janela ATIVA, que podia ser de outro workbook/sheet.
        """
        with self._janela_visivel():
            self._ensure_full_activation()
            janela = self.app.ActiveWindow
            janela.FreezePanes = False
            if unfreeze:
                return True
            if cell:
                self._ws.Range(cell).Select()
            janela.FreezePanes = True
        return True

    # ========== PROTEÇÃO ==========

    def protect_sheet(self, sheet: str = None, password: str = None) -> bool:
        """Protege uma sheet."""
        ws = self._sheet(sheet)
        if password:
            ws.Protect(Password=password)
        else:
            ws.Protect()
        return True

    def unprotect_sheet(self, sheet: str = None, password: str = None) -> bool:
        """Desprotege uma sheet."""
        ws = self._sheet(sheet)
        if password:
            ws.Unprotect(Password=password)
        else:
            ws.Unprotect()
        return True

    # ========== FORMATAR COLUNA/CÉLULA ==========

    def set_column_format(self, column: str, number_format: str) -> bool:
        """Formata tipo da coluna. Ex: xl.set_column_format('A', '#,##0.00')"""
        self._ws.Columns(column).NumberFormat = number_format
        return True

    def set_cell_format(self, cell_range: str, number_format: str) -> bool:
        """Formata tipo da célula. Ex: xl.set_cell_format('A1', 'dd/mm/yyyy')"""
        self._ws.Range(cell_range).NumberFormat = number_format
        return True
