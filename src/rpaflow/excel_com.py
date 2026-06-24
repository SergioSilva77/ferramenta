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

    def auto_fit(self, columns: str = "A:Z") -> bool:
        """Ajusta largura das colunas."""
        self._ws.Columns(columns).AutoFit()
        return True

    # ========== SHEET ==========

    def select_sheet(self, name_or_index) -> bool:
        """Seleciona uma sheet por nome ou índice (1-based)."""
        if isinstance(name_or_index, int):
            self._ws = self._wb.Worksheets(name_or_index)
        else:
            self._ws = self._wb.Worksheets(name_or_index)
        self._ws.Activate()
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

    # ========== TABELA (ListObject) ==========

    def read_table(self, table_name: str) -> list:
        """Lê corpo da tabela (ListObject)."""
        tbl = self._ws.ListObjects(table_name)
        data = tbl.DataBodyRange
        return [[data.Cells(r, c).Value for c in range(1, data.Columns.Count + 1)] for r in range(1, data.Rows.Count + 1)]

    def read_table_header(self, table_name: str) -> list:
        """Lê cabeçalho da tabela."""
        tbl = self._ws.ListObjects(table_name)
        header = tbl.HeaderRowRange
        return [header.Cells(1, c).Value for c in range(1, header.Columns.Count + 1)]

    def read_table_column(self, table_name: str, column_name: str) -> list:
        """Lê uma coluna da tabela pelo nome."""
        tbl = self._ws.ListObjects(table_name)
        col = tbl.ListColumns(column_name).DataBodyRange
        return [col.Cells(r, 1).Value for r in range(1, col.Rows.Count + 1)]

    def count_table_rows(self, table_name: str) -> int:
        """Conta linhas da tabela."""
        tbl = self._ws.ListObjects(table_name)
        return tbl.DataBodyRange.Rows.Count

    def refresh_table(self, table_name: str) -> bool:
        """Atualiza uma tabela."""
        tbl = self._ws.ListObjects(table_name)
        tbl.QueryTable.Refresh(BackgroundQuery=False)
        return True

    def refresh_all(self) -> bool:
        """Atualiza todas as conexões e tabelas."""
        self._wb.RefreshAll()
        return True

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
