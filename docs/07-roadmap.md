# 07 - Roadmap de Implementação

## Fase 1: Core (Implementar primeiro)

| # | Tarefa | Dependências |
|---|--------|--------------|
| 1 | Instalar ANTLR4 + gerar código | — |
| 2 | Criar `grammar/rpa.g4` | #1 |
| 3 | Gerar lexer/parser Python | #1, #2 |
| 4 | Criar `plugins/base.py` (interface) | — |
| 5 | Criar `core/plugin_manager.py` | #4 |
| 6 | Criar `core/scope.py` (variáveis) | — |
| 7 | Criar `core/types.py` (tipos) | — |
| 8 | Criar `core/errors.py` (erros) | — |
| 9 | Criar `core/interpreter.py` (visitor) | #3, #5, #6, #7, #8 |
| 10 | Criar `plugins/core_commands.py` | #4 |
| 11 | Criar `main.py` + `build.py` | #9, #10 |
| 12 | Testar com script exemplo | #11 |

## Fase 2: Plugins (Adicionar conforme necessário)

| # | Tarefa | Dependências |
|---|--------|--------------|
| 13 | Criar `plugins/file_plugin.py` | #4 |
| 14 | Criar `plugins/browser_plugin.py` | #4 |
| 15 | Criar `plugins/excel_plugin.py` | #4 |
| 16 | Criar `plugins/database_plugin.py` | #4 |
| 17 | Criar `plugins/email_plugin.py` | #4 |
| 18 | Criar `plugins/api_plugin.py` | #4 |

## Requisitos

```text
# requirements.txt
antlr4-python3-runtime==4.13.2
```

Plugins adicionais (instalar conforme necessário):
```text
selenium
playwright
openpyxl
pymysql
requests
```
