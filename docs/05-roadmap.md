# 05 - Roadmap

## Fase 1: Estrutura Base

| # | Tarefa |
|---|--------|
| 1 | Criar estrutura do projeto com `pyproject.toml` |
| 2 | Criar `__init__.py` |

## Fase 2: Módulos Core

| # | Tarefa |
|---|--------|
| 3 | Criar `sql.py` (classe SQL) |
| 4 | Criar `files.py` (classe Files) |
| 5 | Criar `api.py` (classe API) |

## Fase 3: Módulos Estendidos

| # | Tarefa |
|---|--------|
| 6 | Criar `excel.py` (classe Excel) |
| 7 | Criar `email.py` (classe Email) |
| 8 | Criar `browser.py` (classe Browser) |

## Fase 4: Validação

| # | Tarefa |
|---|--------|
| 9 | Testar `pip install -e .` |
| 10 | Criar testes unitários |
| 11 | Atualizar README e docs |

## Dependências

| Módulo | Biblioteca | pip install |
|--------|-----------|-------------|
| sql | pymysql | `pip install rpaflow[sql]` |
| excel | openpyxl | `pip install rpaflow[excel]` |
| browser | playwright | `pip install rpaflow[browser]` |
| files | — | `pip install rpaflow[files]` |
| api | httpx | `pip install rpaflow[api]` |
| email | — (stdlib) | `pip install rpaflow[email]` |
