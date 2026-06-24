# rpaflow

Biblioteca Python modular para automação RPA.

## Sumário

| Documento | Descrição |
|-----------|-----------|
| [01 - Visão Geral](docs/01-visao-geral.md) | Objetivo e contexto do projeto |
| [02 - Arquitetura](docs/02-arquitetura.md) | Estrutura do projeto e módulos |
| [03 - Módulos](docs/03-modulos.md) | API completa de cada módulo |
| [04 - Exemplos](docs/04-exemplos.md) | Scripts de exemplo |
| [05 - Roadmap](docs/05-roadmap.md) | Ordem de implementação |

## Módulos

| Módulo | pip install | Descrição |
|--------|-------------|-----------|
| [sql](docs/modulos/sql.md) | `pip install rpaflow[sql]` | MySQL, PostgreSQL, SQL Server, SQLite |
| [excel](docs/modulos/excel.md) | `pip install rpaflow[excel]` | Leitura e escrita de planilhas (openpyxl) |
| [excel_com](docs/modulos/excel_com.md) | `pip install rpaflow[excel-com]` | Automação Excel via COM (Windows) - 60+ métodos |
| [browser](docs/modulos/browser.md) | `pip install rpaflow[browser]` | Automação de navegador |
| [files](docs/modulos/files.md) | `pip install rpaflow[files]` | Operações com arquivos |
| [api](docs/modulos/api.md) | `pip install rpaflow[api]` | Requisições HTTP/REST |
| [email](docs/modulos/email.md) | `pip install rpaflow[email]` | Envio de emails SMTP |

## Quick Start

```bash
pip install rpaflow[excel-com]
```

```python
from rpaflow.excel_com import ExcelCom

xl = ExcelCom(visible=True)
xl.open("C:/dados/vendas.xlsx")

# Filtrar tabela
xl.filter_column_values("Vendas", 1, ["PCD"])       # Mostrar só PCD
xl.filter_column_exclude("Vendas", 1, ["PF", "PJ"]) # Esconder PF e PJ
xl.filter_column_number("Vendas", 3, ">1000")        # Valor > 1000
xl.sort_column("Vendas", 3, order="desc")            # Maior ao menor

# Ler apenas linhas visíveis
dados = xl.read_filtered_table("Vendas")

# Pivot Table
xl.filter_pivot_values("PivotTable1", "Região", ["Este", "Oeste"])

# Sheets ocultas
print(f"Ocultas: {xl.count_hidden_sheets()}")
print(f"Colunas ocultas: {xl.list_hidden_columns()}")

xl.save()
xl.close()
xl.quit()
```

## Instalar Tudo

```bash
pip install rpaflow[all]
```
