# 02 - Arquitetura

## Estrutura do Projeto

```
rpaflow/
├── src/
│   └── rpaflow/
│       ├── __init__.py
│       ├── sql.py              # Classe SQL
│       ├── excel.py            # Classe Excel
│       ├── browser.py          # Classe Browser
│       ├── files.py            # Classe Files
│       ├── api.py              # Classe API
│       └── email.py            # Classe Email
│
├── tests/
├── pyproject.toml
└── README.md
```

## Como Funciona

Cada módulo é uma **classe**. Você importa a classe, instancia e chama os métodos.

```python
from rpaflow.sql import SQL
from rpaflow.excel import Excel

# Conectar no banco
db = SQL(host="localhost", user="root", password="123", database="vendas", type="mysql")
db.connect()
db.insert("clientes", {"nome": "João"})
db.disconnect()

# Planilha
planilha = Excel("dados.xlsx")
planilha.open()
dados = planilha.read("Planilha1", range="A1:D10")
planilha.close()
```

## Dependências

| Módulo | Dependência | pip install |
|--------|-------------|-------------|
| sql | pymysql | `pip install rpaflow[sql]` |
| excel | openpyxl | `pip install rpaflow[excel]` |
| browser | playwright | `pip install rpaflow[browser]` |
| files | — | `pip install rpaflow[files]` |
| api | httpx | `pip install rpaflow[api]` |
| email | — (stdlib) | `pip install rpaflow[email]` |

## Instalação

```bash
pip install rpaflow[sql,excel]
pip install rpaflow[all]
```
