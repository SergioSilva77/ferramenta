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
| [excel_com](docs/modulos/excel_com.md) | `pip install rpaflow[excel-com]` | Automação Excel via COM (Windows) |
| [browser](docs/modulos/browser.md) | `pip install rpaflow[browser]` | Automação de navegador |
| [files](docs/modulos/files.md) | `pip install rpaflow[files]` | Operações com arquivos |
| [api](docs/modulos/api.md) | `pip install rpaflow[api]` | Requisições HTTP/REST |
| [email](docs/modulos/email.md) | `pip install rpaflow[email]` | Envio de emails SMTP |

## Quick Start

```bash
pip install rpaflow[sql]
```

```python
from rpaflow.sql import SQL

db = SQL(type="mysql", host="localhost", user="root", password="123", database="vendas")
db.connect()
db.insert("clientes", {"nome": "João", "email": "joao@email.com"})
db.disconnect()
```

## Instalar Tudo

```bash
pip install rpaflow[all]
```
