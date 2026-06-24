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
| [browser](docs/modulos/browser.md) | `pip install rpaflow[browser]` | Automação de navegador - recursive iframe |
| [desktop](docs/modulos/desktop.md) | `pip install rpaflow[desktop]` | Automação via reconhecimento de imagem |
| [log](docs/modulos/log.md) | `pip install rpaflow[log]` | Logging profissional (inspirado no Serilog) |
| [ini](docs/modulos/ini.md) | `pip install rpaflow[ini]` | Leitura e escrita de arquivos .ini |
| [json](docs/modulos/json.md) | `pip install rpaflow[json]` | Leitura e escrita de JSON com dot notation |
| [files](docs/modulos/files.md) | `pip install rpaflow[files]` | Operações com arquivos |
| [api](docs/modulos/api.md) | `pip install rpaflow[api]` | Requisições HTTP/REST |
| [email](docs/modulos/email.md) | `pip install rpaflow[email]` | Envio de emails SMTP |

## Quick Start - JSON

```bash
pip install rpaflow[json]
```

```python
from rpaflow.json import Json

json = Json()
data = json.load("config.json")

# Dot notation
data.user.name                                   # "Maria"
data.user.address.city                           # "SP"

# Iterar arrays
for bot in data.bots:
    print(f"Bot: {bot.name}, Enabled: {bot.enabled}")

# Iterar dicionário
for empresa, config in data.empresas.items():
    print(f"{empresa}: {config.host}:{config.port}")
```

## Quick Start - INI

```bash
pip install rpaflow[ini]
```

```python
from rpaflow.ini import Ini

ini = Ini("C:/config/robô.ini")

host = ini.get("database", "host")
port = ini.get_int("database", "port")
debug = ini.get_bool("app", "debug")

ini.set("database", "host", "192.168.1.100")
ini.save()
```

## Quick Start - Log

```bash
pip install rpaflow[log]
```

```python
from rpaflow.log import Log

log = Log(path="C:/logs/meu_bot.log", level="DEBUG", json=True)

log = log.bind(bot="vendas", user="admin")
log.info("Iniciando robô")
log.error("Falha na conexão", host="localhost")
log.success("Robô finalizado")
```

## Quick Start - Desktop

```bash
pip install rpaflow[desktop]
```

```python
from rpaflow.desktop import Desktop

desktop = Desktop()

# Localizar e clicar
desktop.click_image("C:/imgs/botao.png")

# Com confiança
desktop.click_image("C:/imgs/botao.png", confidence=0.90)

# Encontrar todas as ocorrências
results = desktop.find_all_images("C:/imgs/icone.png")
```

## Quick Start - Browser

```bash
pip install rpaflow[browser]
```

```python
from rpaflow.browser import Browser

browser = Browser()
browser.start("https://site.com", type="playwright")

# Normal
browser.click("#botao")
browser.type_text("#campo", "texto")

# Recursive (busca em todos os iframes)
browser.click("#botao", recursive=True)

browser.close()
```

## Quick Start - Excel COM

```bash
pip install rpaflow[excel-com]
```

```python
from rpaflow.excel_com import ExcelCom

xl = ExcelCom(visible=True)
xl.open("C:/dados/vendas.xlsx")

xl.filter_column_values("Vendas", 1, ["PCD"])
xl.sort_column("Vendas", 3, order="desc")

xl.save()
xl.close()
xl.quit()
```

## Instalar Tudo

```bash
pip install rpaflow[all]
```
