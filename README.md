# MinhaRPA

Linguagem de programação para automação RPA.

## Sumário

| Documento | Descrição |
|-----------|-----------|
| [01 - Visão Geral](docs/01-visao-geral.md) | Objetivo e contexto do projeto |
| [02 - Arquitetura](docs/02-arquitetura.md) | Estrutura do projeto e fluxo do interpretador |
| [03 - Grammar](docs/03-grammar.md) | Grammar ANTLR completa |
| [04 - Constructos](docs/04-constructos.md) | Sintaxe da linguagem (if, while, for, etc) |
| [05 - Plugins](docs/05-plugins.md) | Sistema de plugins extensíveis |
| [06 - Exemplos](docs/06-exemplos.md) | Scripts de exemplo |
| [07 - Roadmap](docs/07-roadmap.md) | Ordem de implementação |
| [08 - Como Usar](docs/08-uso.md) | Como criar e executar scripts |

## Quick Start

```rpa
log --msg="Olá, mundo!"
```

## Stack

- **ANTLR4** — Gerador de Lexer/Parser
- **Python** — Linguagem do interpretador
- **Plugins** — Módulos extensíveis (browser, excel, banco, etc)
