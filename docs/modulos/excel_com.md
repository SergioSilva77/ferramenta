# Excel COM

Automação do Excel via COM (Windows). Requer Excel instalado.

Diferente do módulo [Excel](excel.md) (openpyxl), o Excel COM pode executar macros, trabalhar com tabelas, atualizar conexões e muito mais.

## Instalar

```bash
pip install rpaflow[excel-com]
```

## Exemplo Rápido

```python
from rpaflow.excel_com import ExcelCom

xl = ExcelCom(visible=True)
xl.open("C:/relatorios/vendas.xlsm")
xl.run_macro("FormatarRelatorio")
xl.save()
xl.close()
xl.quit()
```

## Exemplo Completo

```python
from rpaflow.excel_com import ExcelCom

# Iniciar Excel
xl = ExcelCom(visible=True, display_alerts=False)
xl.open("C:/relatorios/vendas.xlsm")

# Executar macro
xl.run_macro("FormatarRelatorio")

# Executar macro com argumentos
xl.run_macro("ProcessarDados", "arg1", 123)

# Ler valor
valor = xl.get_value("A1")
print(f"Valor: {valor}")

# Ler intervalo
dados = xl.get_value("A1:C10")

# Escrever valor
xl.set_value("D1", "Total")

# Inserir fórmula
xl.set_formula("D2", "=SUM(A2:A10)")

# Inserir fórmula local (pt-BR)
xl.set_formula_local("D2", "=SOMA(A2:A10)")

# Trabalhar com tabelas
tabelas = xl.list_tables()
print(f"Tabelas: {tabelas}")

linhas = xl.count_table_rows("VendasHeader")
print(f"Linhas na tabela: {linhas}")

cabecalho = xl.read_table_header("VendasHeader")
print(f"Colunas: {cabecalho}")

corpo = xl.read_table("VendasHeader")
for linha in corpo:
    print(linha)

# Última linha populada
ultima_planilha = xl.get_last_used_row()
ultima_coluna_a = xl.get_last_used_row_in_column("A")
print(f"Última linha planilha: {ultima_planilha}")
print(f"Última linha coluna A: {ultima_coluna_a}")

# Formatar coluna/célula
xl.set_column_format("B", "#,##0.00")
xl.set_column_format("C", "R$ #,##0.00")
xl.set_cell_format("D1", "dd/mm/yyyy")

# Ocultar/mostrar colunas e linhas
xl.hide_columns("B:D")
xl.show_columns("B:D")
xl.hide_rows(5, 10)
xl.show_rows(5, 10)

# Atualizar dados
xl.refresh_all()
xl.calculate()

# Sheets
sheets = xl.list_sheets()
print(f"Sheets: {sheets}")

xl.select_sheet("Resumo")
xl.add_sheet("NovaSheet")
xl.rename_sheet("NovaSheet", "Dados")
xl.delete_sheet("Dados")

# Proteção
xl.protect_sheet("Resumo", password="123")
xl.unprotect_sheet("Resumo", password="123")

# Performance
xl.set_screen_updating(False)
xl.set_display_alerts(False)

# Window
xl.set_zoom(80)
xl.freeze_panes()

# Info
print(f"Arquivo: {xl.get_filename()}")
print(f"Caminho: {xl.get_filepath()}")
print(f"Última linha: {xl.get_last_row()}")
print(f"Última coluna: {xl.get_last_col()}")

# Salvar e fechar
xl.save()
xl.close()
xl.quit()
```

## Métodos

### Workbook

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `open()` | path, readonly | Abre workbook |
| `save()` | — | Salva workbook |
| `save_as()` | path | Salva como novo arquivo |
| `close()` | save_changes | Fecha workbook |
| `quit()` | — | Fecha instância do Excel |

### Macro

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `run_macro()` | name, *args | Executa macro |

### Célula / Range

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get_value()` | cell_range | Lê valor de célula/intervalo |
| `set_value()` | cell_range, value | Escreve valor |
| `get_cell()` | row, col | Lê valor por linha/coluna (1-based) |
| `set_cell()` | row, col, value | Escreve valor por linha/coluna |
| `select_cell()` | cell_range | Seleciona e ativa célula/intervalo |
| `activate_cell()` | cell_range | Ativa célula (sem selecionar) |
| `set_formula()` | cell_range, formula | Insere fórmula |
| `set_formula_local()` | cell_range, formula | Insere fórmula local (pt-BR) |
| `clear()` | cell_range | Limpa conteúdo |
| `copy_range()` | src_range, dest_range, dest_sheet | Copia intervalo |
| `auto_fit()` | columns | Ajusta largura das colunas |

### Sheet

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `select_sheet()` | name_or_index | Seleciona sheet |
| `list_sheets()` | — | Lista todas as sheets |
| `add_sheet()` | name | Adiciona sheet |
| `rename_sheet()` | old_name, new_name | Renomeia sheet |
| `copy_sheet()` | name, before | Copia sheet |
| `delete_sheet()` | name | Deleta sheet |
| `set_sheet_visible()` | name, visible | Visibilidade da sheet |

### Tabela (ListObject)

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `list_tables()` | — | Lista tabelas da sheet atual |
| `read_table()` | table_name | Lê corpo da tabela |
| `read_table_header()` | table_name | Lê cabeçalho |
| `read_table_column()` | table_name, column_name | Lê coluna |
| `count_table_rows()` | table_name | Conta linhas |
| `refresh_table()` | table_name | Atualiza tabela |
| `refresh_all()` | — | Atualiza tudo |

### Conexões

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `list_connections()` | — | Lista conexões |
| `refresh_connection()` | name | Atualiza conexão |
| `get_connection_string()` | table_name | Lê connection string |
| `get_connection_sql()` | table_name | Lê SQL da conexão |

### Cálculo

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `calculate()` | — | Calcula planilha |
| `wait_calculation()` | — | Aguarda queries assíncronas |

### Info

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get_filename()` | — | Nome do arquivo |
| `get_filepath()` | — | Caminho completo |
| `get_last_row()` | sheet | Última linha usada (UsedRange) |
| `get_last_col()` | sheet | Última coluna usada |
| `get_last_used_row()` | sheet | Última linha com dados (precisa) |
| `get_last_used_row_in_column()` | column, sheet | Última linha populada de uma coluna |

### Ocultar Colunas/Linhas

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `hide_columns()` | columns | Oculta colunas. Ex: `'B:D'` |
| `show_columns()` | columns | Mostra colunas ocultas |
| `hide_rows()` | start_row, end_row | Oculta linhas |
| `show_rows()` | start_row, end_row | Mostra linhas ocultas |

### Formatação

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `set_column_format()` | column, number_format | Formata tipo da coluna |
| `set_cell_format()` | cell_range, number_format | Formata tipo da célula |

#### Formatos Comuns

| Formato | Descrição |
|---------|-----------|
| `#,##0.00` | Número com 2 casas decimais |
| `0%` | Porcentagem |
| `dd/mm/yyyy` | Data |
| `R$ #,##0.00` | Moeda (BRL) |
| `$ #,##0.00` | Moeda (USD) |
| `@` | Texto |
| `0` | Inteiro |

### App / Window

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `set_visible()` | visible | Visibilidade do Excel |
| `set_display_alerts()` | enabled | Exibir alertas |
| `set_screen_updating()` | enabled | Atualização de tela |
| `set_zoom()` | percent | Zoom da janela |
| `freeze_panes()` | — | Congela painéis |

### Proteção

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `protect_sheet()` | sheet, password | Protege sheet |
| `unprotect_sheet()` | sheet, password | Desprotege sheet |
