# 02 - Arquitetura

## Estrutura do Projeto

```
rpaflow/
├── src/
│   └── rpaflow/
│       ├── __init__.py              # Exports públicos
│       ├── _exceptions.py           # Exceções customizadas
│       ├── _deps.py                 # Checker de dependências
│       │
│       ├── sql/                     # Módulo SQL
│       │   ├── __init__.py
│       │   └── _client.py
│       │
│       ├── excel/                   # Módulo Excel
│       │   ├── __init__.py
│       │   └── _reader.py
│       │
│       ├── browser/                 # Módulo Browser
│       │   ├── __init__.py
│       │   └── _driver.py
│       │
│       ├── files/                   # Módulo Arquivos
│       │   ├── __init__.py
│       │   └── _operations.py
│       │
│       ├── api/                     # Módulo HTTP/REST
│       │   ├── __init__.py
│       │   └── _client.py
│       │
│       └── email/                   # Módulo Email
│           ├── __init__.py
│           └── _smtp.py
│
├── tests/
├── pyproject.toml
└── README.md
```

## Dependências por Módulo

| Módulo | Dependência | pip install |
|--------|-------------|-------------|
| core | — | `pip install rpaflow` |
| sql | pymysql | `pip install rpaflow[sql]` |
| excel | openpyxl | `pip install rpaflow[excel]` |
| browser | playwright | `pip install rpaflow[browser]` |
| files | — | `pip install rpaflow[files]` |
| api | httpx | `pip install rpaflow[api]` |
| email | — (stdlib) | `pip install rpaflow[email]` |

## Instalação

```bash
# Só o que precisa
pip install rpaflow[sql,excel]

# Tudo
pip install rpaflow[all]
```
