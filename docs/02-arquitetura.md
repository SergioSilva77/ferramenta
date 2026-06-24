# 02 - Arquitetura

## Princípio Central

A grammar ANTLR é **genérica** — ela não conhece comandos RPA específicos. Apenas parseia `nomeComando argumentos`. O `PluginManager` resolve qual módulo executa cada comando em runtime.

Isso significa: **adicionar um novo comando = criar um arquivo `_plugin.py` na pasta `plugins/`**. Sem tocar na grammar, sem gerar código novamente.

## Fluxo

```
Script .rpa
    │
    ▼
ANTLR Lexer/Parser (grammar/rpa.g4)
    │  → Não sabe o que é "startBrowser"
    │  → Só vê: ID="startBrowser" + args
    ▼
Interpreter (core/interpreter.py)
    │  → visitCmdGenerico()
    │  → command="startBrowser", args={url: "...", type: "playwright"}
    ▼
PluginManager (core/plugin_manager.py)
    │  → Procura "startBrowser" nos plugins carregados
    │  → Encontra em browser_plugin.py
    ▼
Plugin executa a ação
```

## Estrutura de Diretórios

```
rpa-language/
│
├── grammar/
│   └── rpa.g4                        # Grammar genérica (não muda mais)
│
├── generated/                         # Código gerado pelo ANTLR
│   ├── rpaLexer.py
│   ├── rpaParser.py
│   └── rpaVisitor.py
│
├── core/                              # Motor da linguagem
│   ├── __init__.py
│   ├── interpreter.py                 # Visitor que executa comandos
│   ├── scope.py                       # Variáveis ($var) e escopo
│   ├── plugin_manager.py              # Carrega e gerencia plugins
│   ├── types.py                       # Tipos: String, Int, List, Map
│   └── errors.py                      # Erros customizados
│
├── plugins/                           # MÓDULOS EXTENSÍVEIS
│   ├── __init__.py
│   ├── base.py                        # Interface que todo plugin segue
│   ├── core_commands.py               # Sempre carrega: log, wait, sleep
│   ├── file_plugin.py                 # readFile, writeFile, copyFile
│   ├── browser_plugin.py              # startBrowser, click, typeText [FUTURO]
│   ├── excel_plugin.py                # openExcel, readSheet [FUTURO]
│   ├── database_plugin.py             # connectDB, query [FUTURO]
│   ├── email_plugin.py                # sendEmail [FUTURO]
│   └── api_plugin.py                  # httpGet, httpPost [FUTURO]
│
├── examples/
│   ├── hello.rpa
│   └── web_scraping.rpa
│
├── main.py                            # Entry point
├── build.py                           # Gera lexer/parser com ANTLR
└── requirements.txt
```
