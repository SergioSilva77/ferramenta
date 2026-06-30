"""Módulo Excel COM para rpaflow (Windows only)."""

import os


class ExcelCom:
    """Classe para automação do Excel via COM (win32com).

    Requer Windows com Excel instalado.
    """

    def __init__(self, visible: bool = False, display_alerts: bool = False):
        import win32com.client
        self._xl = win32com.client.Dispatch("Excel.Application")
        self._xl.Visible = visible
        self._xl.DisplayAlerts = display_alerts
        self._wb = None
        self._ws = None

    # ========== HELPERS INTERNOS ==========

    def _ensure_full_activation(self):
        """Garante a cadeia completa: App > Workbook > Worksheet."""
        was_visible = self._xl.Visible
        if not was_visible:
            self._xl.Visible = True
        try:
            self._wb.Activate()
            self._ws.Activate()
        except Exception:
            self._xl.Visible = was_visible
            raise

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
        raw = col_data.Value
        if raw is None:
            return []
        if isinstance(raw, tuple):
            if isinstance(raw[0], tuple):
                values = {row[0] for row in raw if row[0] is not None}
            else:
                values = {v for v in raw if v is not None}
        else:
            values = {raw} if raw is not None else set()
        return list(values)

    # ========== WORKBOOK ==========

    def open(self, path: str, readonly: bool = False) -> bool:
        """Abre um workbook."""
        path = os.path.abspath(path)
        self._wb = self._xl.Workbooks.Open(path, ReadOnly=readonly)
        self._ws = self._wb.Worksheets(1)
        return True

    def save(self) -> bool:
        """Salva o workbook atual."""
        self._wb.Save()
        return True

    def save_as(self, path: str) -> bool:
        """Salva o workbook em um novo arquivo."""
        path = os.path.abspath(path)
        self._wb.SaveAs(path)
        return True

    def close(self, save_changes: bool = True) -> bool:
        """Fecha o workbook atual."""
        self._wb.Close(SaveChanges=save_changes)
        self._wb = None
        self._ws = None
        return True

    def quit(self) -> None:
        """Fecha a instância do Excel."""
        try:
            self._xl.Quit()
        except Exception:
            pass
        self._xl = None

    # ========== MACRO ==========

    def run_macro(self, name: str, *args):
        """Executa uma macro. Ex: xl.run_macro('NomeDaMacro', arg1, arg2)"""
        return self._xl.Application.Run(name, *args)

    # ========== CÉLULA / RANGE ==========

    def get_value(self, cell_range: str):
        """Lê valor de célula ou intervalo. Ex: xl.get_value('A1') ou xl.get_value('A1:C10')"""
        return self._ws.Range(cell_range).Value

    def set_value(self, cell_range: str, value) -> bool:
        """Escreve valor em célula ou intervalo."""
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
        self._ensure_full_activation()
        self._ws.Range(cell_range).Select()
        return True

    def activate_cell(self, cell_range: str) -> bool:
        """Ativa uma célula (sem selecionar o intervalo)."""
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

    def copy_range(self, src_range: str, dest_range: str, dest_sheet=None) -> bool:
        """Copia intervalo para outro local."""
        if dest_sheet:
            self._ws.Range(src_range).Copy(dest_sheet.Range(dest_range))
        else:
            self._ws.Range(src_range).Copy(self._ws.Range(dest_range))
        return True

    def copy_values(self, src_range: str, dest_range: str, dest_sheet=None) -> bool:
        """Copia apenas valores (sem fórmulas)."""
        if dest_sheet:
            self._ws.Range(src_range).Copy()
            dest_sheet.Range(dest_range).PasteSpecial(Paste=-4163)  # xlPasteValues
        else:
            self._ws.Range(src_range).Copy()
            self._ws.Range(dest_range).PasteSpecial(Paste=-4163)  # xlPasteValues
        self._xl.CutCopyMode = False
        return True

    def copy_format(self, src_range: str, dest_range: str, dest_sheet=None) -> bool:
        """Copia apenas a formatação."""
        if dest_sheet:
            self._ws.Range(src_range).Copy()
            dest_sheet.Range(dest_range).PasteSpecial(Paste=-4122)  # xlPasteFormats
        else:
            self._ws.Range(src_range).Copy()
            self._ws.Range(dest_range).PasteSpecial(Paste=-4122)  # xlPasteFormats
        self._xl.CutCopyMode = False
        return True

    def move_range(self, src_range: str, dest_range: str, dest_sheet=None) -> bool:
        """Move (recorta e cola) o range para outro local."""
        if dest_sheet:
            self._ws.Range(src_range).Cut(dest_sheet.Range(dest_range))
        else:
            self._ws.Range(src_range).Cut(self._ws.Range(dest_range))
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
        if isinstance(name_or_index, int):
            self._ws = self._wb.Worksheets(name_or_index)
        else:
            self._ws = self._wb.Worksheets(name_or_index)
        self._ensure_full_activation()
        return True

    def list_sheets(self) -> list:
        """Lista todas as sheets do workbook."""
        return [self._wb.Worksheets(i).Name for i in range(1, self._wb.Worksheets.Count + 1)]

    def add_sheet(self, name: str = None) -> bool:
        """Adiciona uma nova sheet."""
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
        """Deleta uma sheet."""
        self._wb.Worksheets(name).Delete()
        return True

    def set_sheet_visible(self, name: str, visible: bool = True) -> bool:
        """Controla visibilidade da sheet. Use False para ocultar, -2 para very hidden."""
        self._wb.Worksheets(name).Visible = visible
        return True

    def count_hidden_sheets(self) -> int:
        """Conta sheets ocultas."""
        count = 0
        for i in range(1, self._wb.Worksheets.Count + 1):
            if not self._wb.Worksheets(i).Visible:
                count += 1
        return count

    def list_hidden_sheets(self) -> list:
        """Lista sheets ocultas."""
        hidden = []
        for i in range(1, self._wb.Worksheets.Count + 1):
            if not self._wb.Worksheets(i).Visible:
                hidden.append(self._wb.Worksheets(i).Name)
        return hidden

    def list_visible_sheets(self) -> list:
        """Lista sheets visíveis."""
        visible = []
        for i in range(1, self._wb.Worksheets.Count + 1):
            if self._wb.Worksheets(i).Visible:
                visible.append(self._wb.Worksheets(i).Name)
        return visible

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
        """Lista colunas ocultas (letras)."""
        hidden = []
        for c in range(1, self._ws.UsedRange.Columns.Count + 1):
            if self._ws.Columns(c).Hidden:
                hidden.append(self._get_column_letter(c))
        return hidden

    def list_hidden_rows(self) -> list:
        """Lista linhas ocultas (números)."""
        hidden = []
        for r in range(1, self._ws.UsedRange.Rows.Count + 1):
            if self._ws.Rows(r).Hidden:
                hidden.append(r)
        return hidden

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
        raw = data.Value
        if raw is None:
            return []
        if not isinstance(raw, tuple):
            return [[raw]]
        return [list(row) if isinstance(row, tuple) else [row] for row in raw]

    def read_filtered_table(self, table_name: str) -> list:
        """Lê apenas linhas visíveis (não filtradas) da tabela — bulk read."""
        tbl = self._ws.ListObjects(table_name)
        data = tbl.DataBodyRange
        if data is None:
            return []

        # BULK: ler tudo de uma vez
        raw = data.Value
        if raw is None:
            return []
        if not isinstance(raw, tuple):
            raw = ((raw,),)

        # BULK: linhas ocultas
        try:
            hidden_flags = data.Rows.Hidden
            if not isinstance(hidden_flags, tuple):
                hidden_flags = (hidden_flags,)
        except Exception:
            hidden_flags = tuple(False for _ in range(len(raw)))

        return [
            [v if v is not None else "" for v in (row if isinstance(row, tuple) else (row,))]
            for i, row in enumerate(raw)
            if i < len(hidden_flags) and not hidden_flags[i]
        ]

    def count_filtered_rows(self, table_name: str) -> int:
        """Conta linhas visíveis (não filtradas) da tabela."""
        return len(self.read_filtered_table(table_name))

    def read_table_header(self, table_name: str) -> list:
        """Lê cabeçalho da tabela — bulk read."""
        tbl = self._ws.ListObjects(table_name)
        header = tbl.HeaderRowRange
        raw = header.Value
        if isinstance(raw, tuple):
            if isinstance(raw[0], tuple):
                return list(raw[0])
            return list(raw)
        return [raw]

    def read_table_column(self, table_name: str, column_name: str) -> list:
        """Lê uma coluna da tabela pelo nome — bulk read."""
        tbl = self._ws.ListObjects(table_name)
        col = tbl.ListColumns(column_name).DataBodyRange
        raw = col.Value
        if raw is None:
            return []
        if isinstance(raw, tuple):
            if isinstance(raw[0], tuple):
                return [row[0] for row in raw]
            return list(raw)
        return [raw]

    def count_table_rows(self, table_name: str) -> int:
        """Conta linhas da tabela."""
        tbl = self._ws.ListObjects(table_name)
        return tbl.DataBodyRange.Rows.Count

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
        if data is None:
            header = self.read_table_header(table_name)
            nomes = [str(h or "").strip() for h in header]
            data = self.read_table(table_name)
        else:
            header = self.read_table_header(table_name)
            nomes = [str(h or "").strip() for h in header]

        for idx, row in enumerate(data):
            if all(
                str(row[nomes.index(k)] or "").strip().replace(".0", "") == str(v).strip().replace(".0", "")
                for k, v in criteria.items() if k in nomes
            ):
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
        ws = self._wb.Worksheets(sheet) if sheet else self._ws
        if start_row is None:
            start_row = header_row + 1

        last_col = ws.UsedRange.Columns.Count
        last_row = ws.UsedRange.Rows.Count

        if start_row > last_row:
            return []

        # BULK: ler cabeçalho de uma vez
        header_raw = ws.Range(ws.Cells(header_row, 1), ws.Cells(header_row, last_col)).Value
        if isinstance(header_raw, tuple):
            if isinstance(header_raw[0], tuple):
                header = [str(h or "").strip() for h in header_raw[0]]
            else:
                header = [str(h or "").strip() for h in header_raw]
        else:
            header = [str(header_raw or "").strip()]

        # BULK: ler todos os dados de uma vez
        data = ws.Range(ws.Cells(start_row, 1), ws.Cells(last_row, last_col)).Value
        if data is None:
            return []
        if not isinstance(data, tuple):
            data = ((data,),)

        # BULK: linhas ocultas
        try:
            hidden_flags = ws.Range(ws.Cells(start_row, 1), ws.Cells(last_row, 1)).Rows.Hidden
            if not isinstance(hidden_flags, tuple):
                hidden_flags = (hidden_flags,)
        except Exception:
            hidden_flags = tuple(False for _ in range(len(data)))

        # Construir índice de colunas do filtro para acesso rápido
        filter_cols = {}
        if filters:
            for col_name in filters:
                if col_name in header:
                    filter_cols[col_name] = header.index(col_name)

        # Filtrar em Python (zero COM)
        result = []
        for i, row in enumerate(data):
            if i < len(hidden_flags) and hidden_flags[i]:
                continue

            if not isinstance(row, tuple):
                row = (row,)

            row_dict = {}
            for c, nome in enumerate(header):
                row_dict[nome] = row[c] if c < len(row) else None

            # Aplicar filtros
            if filters:
                match = True
                for col_name, valor in filters.items():
                    if col_name in filter_cols:
                        cell_val = str(row_dict.get(col_name) or "").strip()
                        filter_val = str(valor or "").strip()
                        if valor is None:
                            if cell_val == "":
                                match = False
                        elif cell_val != filter_val:
                            match = False
                if not match:
                    continue

            result.append(row_dict)

        return result

    def refresh_table(self, table_name: str) -> bool:
        """Atualiza uma tabela."""
        tbl = self._ws.ListObjects(table_name)
        tbl.QueryTable.Refresh(BackgroundQuery=False)
        return True

    def refresh_all(self) -> bool:
        """Atualiza todas as conexões e tabelas."""
        self._wb.RefreshAll()
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
        tbl.Range.AutoFilter(Field=col_idx, Criteria1=values, Operator=7)  # xlFilterValues
        return True

    def filter_column_exclude(self, table_name: str, column, values: list, ignore_case: bool = True) -> bool:
        """Esconde valores específicos."""
        col_idx = self._get_column_index(table_name, column, ignore_case)
        all_values = self._get_column_unique_values(table_name, col_idx)
        visible = [v for v in all_values if v not in values]
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
        tbl = self._ws.ListObjects(table_name)
        try:
            tbl.Range.ShowAllData
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
        ws = self._wb.Worksheets(sheet) if sheet else self._ws
        return [ws.PivotTables(i).Name for i in range(1, ws.PivotTables().Count + 1)]

    def refresh_pivot_table(self, name: str, sheet: str = None) -> bool:
        """Atualiza PivotTable."""
        ws = self._wb.Worksheets(sheet) if sheet else self._ws
        ws.PivotTables(name).RefreshTable()
        return True

    def filter_pivot_values(self, pivot_name: str, field_name: str, visible_items: list, sheet: str = None) -> bool:
        """Filtra itens visíveis da PivotTable."""
        ws = self._wb.Worksheets(sheet) if sheet else self._ws
        pt = ws.PivotTables(pivot_name)
        pf = pt.PivotFields(field_name)
        pf.ClearAllFilters()
        for i in range(1, pf.PivotItems().Count + 1):
            pi = pf.PivotItems(i)
            pi.Visible = pi.Name in visible_items
        return True

    def filter_pivot_olap(self, pivot_name: str, field_name: str, value: str, sheet: str = None) -> bool:
        """Filtra PivotTable OLAP usando CurrentPage.
        
        Args:
            pivot_name: Nome da PivotTable
            field_name: Nome do campo OLAP. Ex: "[Base_atualizada].[TipoLead].[TipoLead]"
            value: Valor para filtrar. Ex: "[Base_atualizada].[TipoLead].&[CONSÓRCIO]"
            sheet: Nome da sheet (opcional)
        """
        ws = self._wb.Worksheets(sheet) if sheet else self._ws
        pt = ws.PivotTables(pivot_name)
        pf = pt.PivotFields(field_name)
        pf.ClearAllFilters()
        pf.CurrentPage = value
        return True

    def filter_pivot_exclude(self, pivot_name: str, field_name: str, exclude_items: list, sheet: str = None) -> bool:
        """Esconde itens específicos da PivotTable."""
        ws = self._wb.Worksheets(sheet) if sheet else self._ws
        pt = ws.PivotTables(pivot_name)
        pf = pt.PivotFields(field_name)
        pf.ClearAllFilters()
        for i in range(1, pf.PivotItems().Count + 1):
            pi = pf.PivotItems(i)
            pi.Visible = pi.Name not in exclude_items
        return True

    def clear_pivot_filters(self, pivot_name: str, sheet: str = None) -> bool:
        """Limpa filtros da PivotTable."""
        ws = self._wb.Worksheets(sheet) if sheet else self._ws
        pt = ws.PivotTables(pivot_name)
        for i in range(1, pt.PivotFields().Count + 1):
            pt.PivotFields(i).ClearAllFilters()
        return True

    def set_pivot_page_filter(self, pivot_name: str, field_name: str, value: str, sheet: str = None) -> bool:
        """Define filtro de página da PivotTable."""
        ws = self._wb.Worksheets(sheet) if sheet else self._ws
        ws.PivotTables(pivot_name).PivotFields(field_name).CurrentPage = value
        return True

    def get_pivot_field_items(self, pivot_name: str, field_name: str, sheet: str = None) -> list:
        """Lista itens de um campo da PivotTable."""
        ws = self._wb.Worksheets(sheet) if sheet else self._ws
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
        self._xl.Calculate()
        return True

    def wait_calculation(self) -> bool:
        """Aguarda queries assíncronas finalizarem."""
        self._xl.CalculateUntilAsyncQueriesStop()
        return True

    # ========== INFO ==========

    def get_filename(self) -> str:
        """Retorna o nome do arquivo."""
        return self._wb.Name

    def get_filepath(self) -> str:
        """Retorna o caminho completo do arquivo."""
        return self._wb.FullName

    def get_last_row(self, sheet: str = None) -> int:
        """Retorna a última linha usada."""
        ws = self._wb.Worksheets(sheet) if sheet else self._ws
        return ws.UsedRange.Rows.Count

    def get_last_col(self, sheet: str = None) -> int:
        """Retorna a última coluna usada."""
        ws = self._wb.Worksheets(sheet) if sheet else self._ws
        return ws.UsedRange.Columns.Count

    def get_last_used_row(self, sheet: str = None) -> int:
        """Retorna a última linha com dados (precisa)."""
        ws = self._wb.Worksheets(sheet) if sheet else self._ws
        return ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

    def get_last_used_row_in_column(self, column: str, sheet: str = None) -> int:
        """Retorna a última linha populada de uma coluna específica."""
        ws = self._wb.Worksheets(sheet) if sheet else self._ws
        return ws.Cells(ws.Rows.Count, column).End(-4162).Row  # xlUp = -4162

    # ========== APP / WINDOW ==========

    def set_visible(self, visible: bool = True) -> bool:
        """Controla visibilidade do Excel."""
        self._xl.Visible = visible
        return True

    def set_display_alerts(self, enabled: bool = True) -> bool:
        """Controla exibição de alertas."""
        self._xl.DisplayAlerts = enabled
        return True

    def set_screen_updating(self, enabled: bool = True) -> bool:
        """Controla atualização de tela (performance)."""
        self._xl.ScreenUpdating = enabled
        return True

    def set_zoom(self, percent: int = 100) -> bool:
        """Define zoom da janela."""
        self._wb.Windows(1).Zoom = percent
        return True

    def freeze_panes(self) -> bool:
        """Congela painéis na posição atual."""
        self._xl.ActiveWindow.FreezePanes = True
        return True

    # ========== PROTEÇÃO ==========

    def protect_sheet(self, sheet: str = None, password: str = None) -> bool:
        """Protege uma sheet."""
        ws = self._wb.Worksheets(sheet) if sheet else self._ws
        if password:
            ws.Protect(Password=password)
        else:
            ws.Protect()
        return True

    def unprotect_sheet(self, sheet: str = None, password: str = None) -> bool:
        """Desprotege uma sheet."""
        ws = self._wb.Worksheets(sheet) if sheet else self._ws
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
