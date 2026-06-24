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

xl = ExcelCom(visible=True, display_alerts=False)
xl.open("C:/dados/vendas.xlsx")

# ====== SHEETS ======
xl.select_sheet("Resumo")
print(f"Sheets: {xl.list_sheets()}")
print(f"Ocultas: {xl.count_hidden_sheets()}")
print(f"Lista ocultas: {xl.list_hidden_sheets()}")
print(f"Visíveis: {xl.list_visible_sheets()}")

# ====== COLUNAS/LINHAS OCULTAS ======
xl.hide_columns("B:D")
xl.hide_rows(5, 10)
print(f"Colunas ocultas: {xl.list_hidden_columns()}")  # ['B', 'C', 'D']
print(f"Linhas ocultas: {xl.list_hidden_rows()}")      # [5, 6, 7, 8, 9, 10]
xl.show_columns("B:D")
xl.show_rows(5, 10)

# ====== CÉLULAS ======
xl.set_value("A1", "Total")
xl.set_formula("A2", "=SUM(B2:B10)")
valor = xl.get_value("A1")
addr = xl.get_cell_address(3, 2)        # "B3"
valor2 = xl.get_cell_by_address("B3")   # valor da célula B3

# ====== COPIAR / MOVER ======
xl.copy_range("A1:B10", "D1")            # Copia tudo (fórmulas, formatação)
xl.copy_values("A1:B10", "D1")           # Copia apenas valores
xl.copy_format("A1:B10", "D1")           # Copia apenas formatação
xl.move_range("A1:B10", "D1")            # Move (recorta e cola)

# ====== MACROS ======
xl.run_macro("FormatarRelatorio")
xl.run_macro("ProcessarDados", "arg1", 123)

# ====== TABELAS ======
tabelas = xl.list_tables()
print(f"Tabelas: {tabelas}")

# Ler tabela completa
dados = xl.read_table("Vendas")

# Ler apenas linhas visíveis (após filtro)
dados_filtrados = xl.read_filtered_table("Vendas")
print(f"Linhas visíveis: {xl.count_filtered_rows('Vendas')}")

# ====== FILTROS DE TABELA ======
# Filtrar: mostrar só PCD
xl.filter_column_values("Vendas", 1, ["PCD"])

# Filtrar: esconder PF e PJ
xl.filter_column_exclude("Vendas", 1, ["PF", "PJ"])

# Filtrar: valor > 1000
xl.filter_column_number("Vendas", 3, ">1000")

# Filtrar: cor amarela
xl.filter_column_color("Vendas", 2, 0xFFFF, "fill")

# Filtrar: não mostrar vazios
xl.filter_column_blanks("Vendas", 4, exclude_empty=True)

# Filtrar: mostrar apenas vazios
xl.filter_column_blanks("Vendas", 4, exclude_empty=False)

# Filtrar: Status = "Aprovado"
xl.filter_column("Vendas", 2, "Aprovado")

# Limpar filtros
xl.clear_filters("Vendas")

# ====== ORDENAÇÃO ======
# Do maior ao menor
xl.sort_column("Vendas", 3, order="desc")

# Do menor ao maior
xl.sort_column("Vendas", 3, order="asc")

# ====== TABELA DINÂMICA (PivotTable) ======
pivots = xl.list_pivot_tables()
print(f"PivotTables: {pivots}")

# Atualizar
xl.refresh_pivot_table("PivotTable1")

# Filtrar: mostrar só Este e Oeste
xl.filter_pivot_values("PivotTable1", "Região", ["Este", "Oeste"])

# Filtrar PivotTable OLAP
xl.filter_pivot_olap(
    "Tabela dinâmica9",
    "[Base_atualizada].[TipoLead].[TipoLead]",
    "[Base_atualizada].[TipoLead].&[CONSÓRCIO]"
)

# Filtrar: esconder Norte e Sul
xl.filter_pivot_exclude("PivotTable1", "Região", ["Norte", "Sul"])

# Limpar filtros pivot
xl.clear_pivot_filters("PivotTable1")

# Filtro de página
xl.set_pivot_page_filter("PivotTable1", "Ano", "2025")

# Listar itens de um campo
itens = xl.get_pivot_field_items("PivotTable1", "Região")
# [{'name': 'Este', 'visible': True}, {'name': 'Oeste', 'visible': False}]

# ====== CONEXÕES ======
conexoes = xl.list_connections()
xl.refresh_connection("MinhaConexao")

# ====== CÁLCULO ======
xl.calculate()
xl.wait_calculation()

# ====== INFO ======
print(f"Arquivo: {xl.get_filename()}")
print(f"Caminho: {xl.get_filepath()}")
print(f"Última linha: {xl.get_last_row()}")
print(f"Última coluna: {xl.get_last_col()}")
print(f"Última linha usada: {xl.get_last_used_row()}")
print(f"Última linha coluna A: {xl.get_last_used_row_in_column('A')}")

# ====== PERFORMANCE ======
xl.set_screen_updating(False)
xl.set_display_alerts(False)

# ====== WINDOW ======
xl.set_zoom(80)
xl.freeze_panes()

# ====== PROTEÇÃO ======
xl.protect_sheet("Resumo", password="123")
xl.unprotect_sheet("Resumo", password="123")

# ====== FORMATAÇÃO ======
xl.set_column_format("B", "#,##0.00")
xl.set_column_format("C", "R$ #,##0.00")
xl.set_cell_format("D1", "dd/mm/yyyy")

# ====== SALVAR E FECHAR ======
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
| `copy_values()` | src_range, dest_range, dest_sheet | Copia apenas valores |
| `copy_format()` | src_range, dest_range, dest_sheet | Copia apenas formatação |
| `move_range()` | src_range, dest_range, dest_sheet | Move (recorta e cola) |
| `auto_fit()` | columns | Ajusta largura das colunas |
| `get_cell_address()` | row, col | Retorna endereço. Ex: (1,1) -> "A1" |
| `get_cell_by_address()` | address | Lê valor por endereço. Ex: "A1" |

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
| `count_hidden_sheets()` | — | Conta sheets ocultas |
| `list_hidden_sheets()` | — | Lista sheets ocultas |
| `list_visible_sheets()` | — | Lista sheets visíveis |

### Ocultar Colunas/Linhas

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `hide_columns()` | columns | Oculta colunas. Ex: `'B:D'` |
| `show_columns()` | columns | Mostra colunas ocultas |
| `hide_rows()` | start_row, end_row | Oculta linhas |
| `show_rows()` | start_row, end_row | Mostra linhas ocultas |
| `list_hidden_columns()` | — | Lista colunas ocultas (letras) |
| `list_hidden_rows()` | — | Lista linhas ocultas (números) |

### Tabela (ListObject)

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `list_tables()` | — | Lista tabelas da sheet atual |
| `read_table()` | table_name | Lê corpo da tabela |
| `read_filtered_table()` | table_name | Lê apenas linhas visíveis |
| `count_filtered_rows()` | table_name | Conta linhas visíveis |
| `read_table_header()` | table_name | Lê cabeçalho |
| `read_table_column()` | table_name, column_name | Lê coluna |
| `count_table_rows()` | table_name | Conta linhas |
| `refresh_table()` | table_name | Atualiza tabela |
| `refresh_all()` | — | Atualiza tudo |

### Filtros de Tabela

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `filter_column()` | table, column, criteria | Filtro simples (=, <>, contém, etc.) |
| `filter_column_values()` | table, column, values | Mostrar apenas valores da lista |
| `filter_column_exclude()` | table, column, values | Esconder valores da lista |
| `filter_column_number()` | table, column, criteria | Filtro numérico (>, <, >=, <=) |
| `filter_column_color()` | table, column, color, type | Filtrar por cor (fill/font) |
| `filter_column_blanks()` | table, column, exclude_empty | Filtrar vazios/não-vazios |
| `clear_filters()` | table | Limpa todos os filtros |
| `sort_column()` | table, column, order | Classificar (asc/desc) |

#### Operadores de Filtro

| Operador | Sintaxe | Exemplo |
|----------|---------|---------|
| Igual | `"valor"` | `"Aprovado"` |
| Não igual | `"<>valor"` | `"<>Aprovado"` |
| Contém | `"*texto*"` | `"*Silva*"` |
| Não contém | `"<>*texto*"` | `"<>*Silva*"` |
| Começa com | `"texto*"` | `"Silva*"` |
| Termina com | `"*texto"` | `"*Silva"` |
| Maior que | `">100"` | `">1000"` |
| Menor que | `"<100"` | `"<500"` |
| Maior ou igual | `">=100"` | `">=1000"` |
| Menor ou igual | `"<=100"` | `"<=500"` |
| Vazio | `"="` | `"="` |
| Não vazio | `"<>"` | `"<>"` |

### Tabela Dinâmica (PivotTable)

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `list_pivot_tables()` | sheet | Lista PivotTables |
| `refresh_pivot_table()` | name, sheet | Atualiza PivotTable |
| `filter_pivot_values()` | pivot, field, visible_items | Filtra itens visíveis |
| `filter_pivot_olap()` | pivot, field, value | Filtra PivotTable OLAP (CurrentPage) |
| `filter_pivot_exclude()` | pivot, field, exclude_items | Esconde itens específicos |
| `clear_pivot_filters()` | pivot, sheet | Limpa filtros pivot |
| `set_pivot_page_filter()` | pivot, field, value | Filtro de página |
| `get_pivot_field_items()` | pivot, field | Lista itens do campo |

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
