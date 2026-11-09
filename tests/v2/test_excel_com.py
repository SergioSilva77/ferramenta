"""ExcelCom v2 — integração real com o Excel (Windows). Pulado se não houver Excel instalado."""

import subprocess
import sys

import pytest

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="Excel COM só no Windows")


def _tem_excel() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application"))
        return True
    except OSError:
        return False


if not _tem_excel():
    pytestmark = pytest.mark.skip(reason="Excel não instalado")

from rpaflow.v2.excel_com import ExcelCom, ExcelComError  # noqa: E402


def _pids_excel() -> set:
    saida = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"], capture_output=True, text=True
    ).stdout
    return {linha.split('","')[1] for linha in saida.splitlines() if "EXCEL.EXE" in linha.upper()}


@pytest.fixture
def xl():
    with ExcelCom() as x:
        yield x


@pytest.fixture
def xl_tabela(xl, tmp_path):
    """Workbook com a tabela `T` (NF, Valor, Status) e uma sheet Resumo."""
    caminho = tmp_path / "base.xlsx"
    import win32com.client  # noqa: F401
    wb = xl.app.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "Dados"
    linhas = [("NF", "Valor", "Status"), ("1001", 10.5, "A Pagar"), ("1002", 20, "Pago"), ("10.05", 30, "A Pagar"),
              (1003.0, 40, "A Pagar")]
    for r, linha in enumerate(linhas, 1):
        for c, v in enumerate(linha, 1):
            ws.Cells(r, c).Value = v
    lo = ws.ListObjects.Add(1, ws.Range("A1:C5"), None, 1)
    lo.Name = "T"
    wb.SaveAs(str(caminho), FileFormat=51)
    wb.Close()
    xl.open(str(caminho))
    return xl


def test_cada_instancia_tem_excel_proprio_e_quit_fecha_so_o_dele():
    antes = _pids_excel()
    a, b = ExcelCom(), ExcelCom()
    novos = _pids_excel() - antes
    assert len(novos) == 2  # DispatchEx: dois processos, não um compartilhado
    a.quit()
    a.quit()  # idempotente
    restante = _pids_excel() - antes
    assert len(restante) == 1  # o do `b` continua vivo
    b.quit()
    assert _pids_excel() == antes  # nada sobrou


def test_with_e_del_nao_deixam_excel_aberto():
    antes = _pids_excel()
    with ExcelCom():
        assert len(_pids_excel() - antes) == 1
    x = ExcelCom()
    del x
    assert _pids_excel() == antes


def test_sem_workbook_erro_claro_e_close_inofensivo(xl):
    assert xl.close() is False
    with pytest.raises(ExcelComError, match="open"):
        xl.get_value("A1")


def test_open_arquivo_inexistente(xl, tmp_path):
    with pytest.raises(FileNotFoundError):
        xl.open(str(tmp_path / "nao_existe.xlsx"))


def test_leitura_de_tabela(xl_tabela):
    xl = xl_tabela
    assert xl.list_tables() == ["T"]
    assert xl.read_table_header("T") == ["NF", "Valor", "Status"]
    assert xl.count_table_rows("T") == 4
    assert xl.read_table_column("T", "status")[0] == "A Pagar"  # ignora case
    assert xl.read_table_column("T", 2) == [10.5, 20.0, 30.0, 40.0]
    assert xl.read_table_as_dicts("T")[1]["NF"] == 1002.0  # o Excel converteu "1002" em número ao gravar


def test_find_row_nao_corrompe_valores_e_valida_colunas(xl_tabela):
    xl = xl_tabela
    assert xl.find_row("T", {"NF": "10.05"}) == 2  # v1 transformava "10.05" em "105"
    assert xl.find_row("T", {"NF": "1003"}) == 3  # 1003.0 (float) == "1003"
    assert xl.find_row("T", {"NF": "1002", "Status": "Pago"}) == 1
    assert xl.find_row("T", {"NF": "9999"}) is None
    with pytest.raises(KeyError, match="Colun"):  # v1 devolvia a linha 0
        xl.find_row("T", {"Coluna Que Nao Existe": "x"})


def test_filtro_e_clear_filters_de_verdade(xl_tabela):
    xl = xl_tabela
    xl.filter_column("T", "Status", "A Pagar")
    assert xl.count_filtered_rows("T") == 3
    assert len(xl.read_filtered_table_as_dicts("T")) == 3
    xl.clear_filters("T")  # no v1 isto não limpava nada
    assert xl.count_filtered_rows("T") == 4

    xl.filter_column_values("T", "Valor", [30, 40])  # números: v1 estourava
    assert xl.count_filtered_rows("T") == 2
    xl.clear_filters("T")

    xl.filter_column_exclude("T", "Status", ["Pago"])
    assert xl.count_filtered_rows("T") == 3
    xl.clear_filters("T")


def test_tabela_vazia_nao_estoura(xl_tabela):
    xl = xl_tabela
    xl.clear_filters("T")
    xl.app.DisplayAlerts = False
    lo = xl.sheet.ListObjects("T")
    lo.DataBodyRange.Delete()
    assert xl.count_table_rows("T") == 0
    assert xl.read_table("T") == []
    assert xl.read_table_column("T", "NF") == []
    assert xl.read_filtered_table("T") == []


def test_read_range_filtered_com_used_range_deslocado(xl):
    xl.app.Workbooks.Add()
    xl._wb_obj = xl.app.ActiveWorkbook
    xl._ws_obj = xl._wb_obj.Worksheets(1)
    # tudo começa na linha 5, coluna B: UsedRange.Rows.Count (v1) não é a última linha
    dados = [("Nome", "Status"), ("Ana", "ok"), ("Bia", "falha"), ("Cris", "ok")]
    for r, linha in enumerate(dados, 5):
        for c, v in enumerate(linha, 2):
            xl.set_cell(r, c, v)
    assert xl.get_last_row() == 8
    linhas = xl.read_range_filtered(header_row=5)
    # a coluna A está vazia, então o 1º cabeçalho é "" e os dados vêm por nome
    assert [l["Nome"] for l in linhas] == ["Ana", "Bia", "Cris"]
    assert [l["Nome"] for l in xl.read_range_filtered(5, filters={"Status": "ok"})] == ["Ana", "Cris"]
    with pytest.raises(KeyError):
        xl.read_range_filtered(5, filters={"Nao Existe": "x"})
    # linha oculta no meio (misto -> COM devolve None; v1 tratava como "nenhuma oculta")
    xl.hide_rows(7)
    assert [l["Nome"] for l in xl.read_range_filtered(5)] == ["Ana", "Cris"]
    assert xl.list_hidden_rows() == [7]


def test_sheets_very_hidden_contam_como_ocultas(xl_tabela):
    xl = xl_tabela
    xl.add_sheet("Oculta1", at_end=True)
    xl.add_sheet("Oculta2", at_end=True)
    xl.set_sheet_visible("Oculta1", False)
    xl.set_sheet_visible("Oculta2", "very_hidden")  # v1: -2 (inválido)
    assert sorted(xl.list_hidden_sheets()) == ["Oculta1", "Oculta2"]
    assert xl.count_hidden_sheets() == 2
    assert xl.list_visible_sheets()[0] == "Dados"


def test_set_value_bloco_e_wait_calculation(xl_tabela):
    xl = xl_tabela
    xl.set_value("E1:F2", [[1, 2], [3, 4]])
    xl.set_value("E4:G4", [7, 8, 9])  # lista simples = uma linha
    assert xl.get_value("E1:F2") == ((1.0, 2.0), (3.0, 4.0))
    assert xl.get_value("E4:G4") == ((7.0, 8.0, 9.0),)
    assert xl.wait_calculation() is True  # v1: método COM inexistente
    assert xl.refresh_all() is True


def test_select_nao_deixa_excel_visivel(xl_tabela):
    xl = xl_tabela
    assert xl.app.Visible is False
    xl.select_cell("B2")
    assert xl.app.Visible is False  # v1 deixava visível pra sempre
    xl.freeze_panes("B2")
    xl.freeze_panes(unfreeze=True)
    assert xl.app.Visible is False


def test_save_as_pelo_formato_da_extensao(xl_tabela, tmp_path):
    xl = xl_tabela
    xl.save_as(str(tmp_path / "saida" / "copia.csv"))  # pasta nova + csv de verdade
    texto = (tmp_path / "saida" / "copia.csv").read_text(encoding="cp1252", errors="replace")
    assert "NF" in texto and "1001" in texto


def test_copy_values_entre_sheets_por_nome(xl_tabela):
    xl = xl_tabela
    xl.add_sheet("Destino", at_end=True)
    xl.copy_values("A1:C2", "A1", dest_sheet="Destino")
    assert xl.app.CutCopyMode is False or xl.app.CutCopyMode == 0
    assert xl.sheet.Name == "Dados"
    assert xl.workbook.Worksheets("Destino").Range("B2").Value == 10.5


def test_pivot_troca_fonte_para_tabela(xl_tabela):
    xl = xl_tabela
    wb = xl.workbook
    xl.add_sheet("Resumo", at_end=True)
    ws_res = wb.Worksheets("Resumo")
    cache = wb.PivotCaches().Create(1, "Dados!A1:C3")  # fonte fixa e curta (o bug da Hyundai)
    pt = cache.CreatePivotTable(ws_res.Range("A3"), "PT")
    pt.PivotFields("Status").Orientation = 1
    pt.AddDataField(pt.PivotFields("Valor"), "Soma", -4157)
    xl.select_sheet("Dados")
    xl.set_pivot_source("PT", "T", sheet="Resumo")  # nome da tabela: cresce com os dados
    total = sum(ws_res.Cells(r, 2).Value or 0 for r in range(4, 8) if ws_res.Cells(r, 1).Value not in (None, "Total Geral"))
    assert total == 10.5 + 20 + 30 + 40
    assert xl.refresh_pivot_caches() is True
    assert xl.list_pivot_tables("Resumo") == ["PT"]
