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
| [06 - Plugins](docs/06-plugins.md) | Como criar novos módulos |

## Quick Start

```bash
pip install rpaflow
```

```python
from rpaflow import sql

sql.connect(host="localhost", user="root", password="123", database="vendas", type="mysql")
sql.insert("clientes", {"nome": "João", "email": "joao@email.com"})
sql.disconnect()
```

## Módulos

| Módulo | pip install | Descrição |
|--------|-------------|-----------|
| `sql` | `pip install rpaflow[sql]` | Conectar, consultar, inserir, atualizar, deletar |
| `excel` | `pip install rpaflow[excel]` | Abrir, ler, escrever, salvar planilhas |
| `browser` | `pip install rpaflow[browser]` | Iniciar navegador, clicar, digitar, screenshot |
| `files` | `pip install rpaflow[files]` | Ler, escrever, copiar, mover, deletar arquivos |
| `api` | `pip install rpaflow[api]` | GET, POST, PUT, DELETE em APIs REST |
| `email` | `pip install rpaflow[email]` | Enviar e ler emails |

## Instalar Tudo

```bash
pip install rpaflow[all]
```
