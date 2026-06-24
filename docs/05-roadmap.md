# 05 - Roadmap

## Fase 1: Estrutura Base

| # | Tarefa | Dependências |
|---|--------|--------------|
| 1 | Criar estrutura do projeto | — |
| 2 | Criar `pyproject.toml` | #1 |
| 3 | Criar `_exceptions.py` | — |
| 4 | Criar `_deps.py` | — |
| 5 | Criar `__init__.py` principal | #3, #4 |

## Fase 2: Módulos Core

| # | Tarefa | Dependências |
|---|--------|--------------|
| 6 | Criar módulo `sql` | #5 |
| 7 | Criar módulo `files` | #5 |
| 8 | Criar módulo `api` | #5 |

## Fase 3: Módulos Estendidos

| # | Tarefa | Dependências |
|---|--------|--------------|
| 9 | Criar módulo `excel` | #5 |
| 10 | Criar módulo `email` | #5 |
| 11 | Criar módulo `browser` | #5 |

## Fase 4: Validação

| # | Tarefa | Dependências |
|---|--------|--------------|
| 12 | Testar `pip install -e .` | #6-#11 |
| 13 | Criar testes unitários | #6-#11 |
| 14 | Atualizar README e docs | #12 |

## Fase 5: Publicação

| # | Tarefa | Dependências |
|---|--------|--------------|
| 15 | Build do pacote | #14 |
| 16 | Publicar no PyPI | #15 |

## Dependências dos Módulos

| Módulo | Biblioteca | pip install |
|--------|-----------|-------------|
| sql | pymysql | `pip install rpaflow[sql]` |
| excel | openpyxl | `pip install rpaflow[excel]` |
| browser | playwright | `pip install rpaflow[browser]` |
| files | — | `pip install rpaflow[files]` |
| api | httpx | `pip install rpaflow[api]` |
| email | — (stdlib) | `pip install rpaflow[email]` |
