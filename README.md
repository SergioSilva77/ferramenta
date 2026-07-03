# rpaflow

Biblioteca Python modular para automação RPA.

## Instalação

```bash
pip install rpaflow[all]  # Todos os módulos
pip install rpaflow[sql]  # Módulo específico
```

## Módulos

| Módulo | pip install | Descrição |
|--------|-------------|-----------|
| [sql](#sql) | `pip install rpaflow[sql]` | MySQL, PostgreSQL, SQL Server, SQLite |
| [excel](#excel) | `pip install rpaflow[excel]` | Leitura e escrita de planilhas (openpyxl) |
| [excel_com](#excel_com) | `pip install rpaflow[excel-com]` | Automação Excel via COM (Windows) - 60+ métodos |
| [browser](#browser) | `pip install rpaflow[browser]` | Automação de navegador - recursive iframe |
| [desktop](#desktop) | `pip install rpaflow[desktop]` | Automação via reconhecimento de imagem |
| [log](#log) | `pip install rpaflow[log]` | Logging profissional (inspirado no Serilog) |
| [ini](#ini) | `pip install rpaflow[ini]` | Leitura e escrita de arquivos .ini |
| [json](#json) | `pip install rpaflow[json]` | Leitura e escrita de JSON com dot notation |
| [files](#files) | `pip install rpaflow[files]` | Operações com arquivos |
| [api](#api) | `pip install rpaflow[api]` | Requisições HTTP/REST |
| [email](#email) | `pip install rpaflow[email]` | Envio de emails SMTP |

---

# SQL

Conexão com bancos de dados MySQL, PostgreSQL, SQL Server e SQLite. Inspirado no `System.Data.SqlClient` do C#.

## Instalar

```bash
pip install rpaflow[sql]
```

## Como Usar (Passo a Passo)

```python
from rpaflow.sql import SQL

# Conectar
db = SQL(type="sqlite", database="C:/dados/app.db")
db.connect()

# CREATE
db.execute("CREATE TABLE IF NOT EXISTS usuarios (id INTEGER PRIMARY KEY, nome TEXT, email TEXT)")

# INSERT
db.execute("INSERT INTO usuarios (nome, email) VALUES (?, ?)", "João", "joao@email.com")
db.commit()

# SELECT
usuarios = db.select("usuarios")
joao = db.select("usuarios", where={"nome": "João"})
db.disconnect()
```

## Métodos

### Conexão

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `connect()` | — | Abre conexão |
| `disconnect()` | — | Fecha conexão |
| `is_connected()` | — | True/False |

### Query

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `execute()` | sql, params | Executa SQL (INSERT, UPDATE, DELETE) |
| `commit()` | — | Salva alterações |
| `rollback()` | — | Desfaz alterações |

### Select

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `select()` | table, columns, where, order_by, limit | Retorna lista de dicts |
| `count()` | table, where | Conta registros |
| `exists()` | table, where | True/False |
| `scalar()` | sql, params | Retorna valor único |

### Parâmetros Comuns

| Parâmetro | Padrão | Descrição |
|-----------|--------|-----------|
| `type` | — | mysql, postgresql, sqlite |
| `host` | localhost | Servidor |
| `port` | 5432/3306 | Porta |
| `user` | — | Usuário |
| `password` | — | Senha |
| `database` | — | Nome do banco |

### Parâmetros de Select

| Parâmetro | Tipo | Descrição |
|-----------|------|-----------|
| `columns` | list | Colunas específicas |
| `where` | dict | Filtros |
| `order_by` | str | Ordenação |
| `limit` | int | Limite de registros |

## Exemplos Práticos

### CRUD Completo

```python
from rpaflow.sql import SQL

db = SQL(type="mysql", host="localhost", user="root", password="123", database="vendas")
db.connect()

# CREATE TABLE
db.execute("""
    CREATE TABLE IF NOT EXISTS pedidos (
        id INTEGER PRIMARY KEY AUTO_INCREMENT,
        cliente TEXT,
        valor DECIMAL(10,2),
        status TEXT DEFAULT 'pendente'
    )
""")

# INSERT
db.execute("INSERT INTO pedidos (cliente, valor) VALUES (?, ?)", "Maria", 1500.00)
db.commit()

# SELECT
pendentes = db.select("pedidos", where={"status": "pendente"})
print(f"Pendentes: {len(pendentes)}")

# UPDATE
db.execute("UPDATE pedidos SET status = 'cancelado' WHERE id = ?", 1)
db.commit()

# DELETE
db.execute("DELETE FROM pedidos WHERE id = ?", 1)
db.commit()

db.disconnect()
```

### Usando where (dict)

```python
usuarios = db.select("usuarios", where={"ativo": 1})
print(usuarios)

usuarios = db.select("usuarios", where={"ativo": 1, "nivel": "admin"})
print(usuarios)
```

### Parâmetros nomeados

```python
usuarios = db.select("usuarios", where={"nome": "João"})
pedidos = db.select("pedidos", where={"usuario_id": 1, "status": "pendente"})
```

### SELECT com joins

```python
results = db.select(
    "pedidos",
    columns=["pedidos.id", "clientes.nome", "pedidos.valor"],
    where={"pedidos.status": "pendente"},
    order_by="pedidos.data DESC"
)
```

### Executar SQL raw

```python
db.execute("UPDATE usuarios SET ativo = 0 WHERE data_cadastro < '2024-01-01'")
db.commit()

usuarios = db.select("usuarios", where={"ativo": 0})
```

### Verificar existência

```python
if db.exists("usuarios", {"email": "admin@email.com"}):
    print("Usuário já existe")

total = db.count("pedidos", where={"status": "concluido"})
print(f"Concluídos: {total}")
```

### transactions

```python
db.connect()
try:
    db.execute("UPDATE contas SET saldo = saldo - 100 WHERE id = 1")
    db.execute("UPDATE contas SET saldo = saldo + 100 WHERE id = 2")
    db.commit()
except:
    db.rollback()
```

### Diferentes bancos

```python
# MySQL
db = SQL(type="mysql", host="localhost", user="root", password="123", database="vendas")

# PostgreSQL
db = SQL(type="postgresql", host="localhost", user="postgres", password="123", database="vendas")

# SQL Server
db = SQL(type="sqlserver", host="localhost", user="sa", password="123", database="vendas")

# SQLite
db = SQL(type="sqlite", database="C:/dados/app.db")
```

---

# Excel

Leitura e escrita de planilhas Excel (.xlsx) usando openpyxl. Funciona em qualquer SO.

## Instalar

```bash
pip install rpaflow[excel]
```

## Como Usar

```python
from rpaflow.excel import Excel

excel = Excel()
excel.open("C:/dados/planilha.xlsx")

# Ler dados
celula = excel.get_cell("A1")
valor = excel.get_value("B2")
excel.set_value("C3", "novo valor")

# Salvar
excel.save()
excel.close()
```

## Métodos

### Abertura

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `open()` | filepath, read_only, write_only | Abre planilha |
| `save()` | filepath | Salva planilha |
| `close()` | — | Fecha planilha |
| `create_sheet()` | name | Cria nova aba |

### Leitura

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get_value()` | cell | Retorna valor da célula |
| `get_row()` | row | Retorna linha como tupla |
| `get_column()` | col | Retorna coluna como tupla |
| `get_used_range()` | — | Retorna intervalo usado |
| `get_max_row()` | — | Última linha com dados |
| `get_max_column()` | — | Última coluna com dados |

### Escrita

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `set_value()` | cell, value | Define valor |
| `write_row()` | row, data | Escreve linha |
| `write_column()` | col, data | Escreve coluna |

### Conversão

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `to_list()` | sheet_name | Converte para lista de dicts |
| `to_dict()` | sheet_name, key_column | Converte para dict |
| `from_list()` | data, start_cell | Escreve lista de dicts |
| `from_dict()` | data, start_cell | Escreve dict |

## Parâmetros Comuns

| Parâmetro | Padrão | Descrição |
|-----------|--------|-----------|
| `filepath` | — | Caminho do arquivo |
| `cell` | — | Referência (ex: "A1") |
| `sheet_name` | Ativa | Nome da aba |
| `row` | — | Número da linha |
| `col` | — | Número da coluna |

---

# ExcelCom

Automação do Excel via COM (Windows). Controle completo: 60+ métodos para ler, escrever, filtrar, ordenar, formatar, tabelas dinâmicas, proteção de planilhas e formatação condicional.

## Instalar

```bash
pip install rpaflow[excel-com]
```

## Requisitos

- Windows
- Microsoft Excel instalado

## Como Usar (Passo a Passo)

```python
from rpaflow.excel_com import ExcelCom

xl = ExcelCom(visible=True)
xl.open("C:/dados/vendas.xlsx")

# Ler valor
valor = xl.get_value("A1")

# Escrever
xl.set_value("B1", "Nome")

# Filtrar por valores específicos
xl.filter_column_values("Vendas", 1, ["PCD"])

# Ordenar
xl.sort_column("Vendas", 3, order="desc")

xl.save()
xl.close()
xl.quit()
```

## Métodos

### Abertura/Busca

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `open()` | filepath | Abre arquivo Excel |
| `open_recent(days)` | days | Abre arquivo usado recentemente |
| `find_files(directory)` | directory | Lista arquivos Excel |
| `find_files_recursive(directory)` | directory | Lista recursivamente |
| `check_open()` | — | Verifica se Excel está aberto |

### Atalhos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `new_workbook()` | — | Cria nova pasta de trabalho |
| `save()` | — | Salva arquivo |
| `save_as()` | filepath | Salva como novo arquivo |
| `close()` | — | Fecha pasta de trabalho |
| `quit()` | — | Fecha Excel |
| `close_without_saving()` | — | Fecha sem salvar |

### Planilha

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get_workbook_name()` | — | Nome do arquivo |
| `select_sheet()` | sheet_name | Seleciona aba |
| `rename_sheet()` | old_name, new_name | Renomeia aba |
| `copy_sheet()` | sheet_name | Copia aba |
| `delete_sheet()` | sheet_name | Deleta aba |
| `get_sheets_names()` | — | Lista todas as abas |
| `get_sheets_count()` | — | Conta abas |
| `add_worksheet()` | name | Adiciona nova aba |

### Célula

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get_value()` | cell | Retorna valor |
| `set_value()` | cell, value | Define valor |
| `get_values()` | range | Retorna intervalo |
| `set_values()` | range, values | Define intervalo |
| `clear_cell()` | cell | Limpa célula |
| `clear_range()` | range | Limpa intervalo |
| `select_cell()` | cell | Seleciona célula |
| `activate_cell()` | cell | Ativa e foca célula |
| `go_to_cell()` | cell | Navega até célula |
| `scroll_to_cell()` | cell | Rola até célula |
| `get_current_cell()` | — | Célula atual |

### Coluna

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `select_column()` | col_num | Seleciona coluna |
| `get_column_width()` | col_num | Largura da coluna |
| `set_column_width()` | col_num, width | Define largura |
| `auto_fit_column()` | col_num | Auto-ajusta largura |
| `get_column_letter()` | col_num | Converte número para letra |
| `get_column_number()` | col_letter | Converte letra para número |

### Linha

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `select_row()` | row_num | Seleciona linha |
| `get_row_height()` | row_num | Altura da linha |
| `set_row_height()` | row_num, height | Define altura |
| `auto_fit_row()` | row_num | Auto-ajusta altura |
| `insert_row()` | row_num, shift | Insere linha |
| `delete_row()` | row_num, shift | Deleta linha |

### Intervalo

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get_used_range()` | — | Intervalo usado |
| `get_used_rows()` | — | Número de linhas usadas |
| `get_used_columns()` | — | Número de colunas usadas |
| `copy_range()` | range_from, range_to | Copia intervalo |
| `cut_range()` | range_from, range_to | Move intervalo |
| `paste_range()` | range | Cola intervalo |

### Tabela Dinâmica

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `create_pivot_table()` | source, destination, name | Cria tabela dinâmica |
| `add_pivot_field()` | pivot_name, field, orientation, position | Adiciona campo |
| `refresh_pivot()` | pivot_name | Atualiza tabela dinâmica |

### Filtro e Ordenação

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `filter_column()` | sheet, col_num, criteria | Filtra por critérios |
| `filter_column_values()` | sheet, col_num, values | Filtra por valores específicos |
| `remove_column_filter()` | sheet, col_num | Remove filtro de coluna |
| `clear_filters()` | sheet | Remove todos os filtros |
| `sort_column()` | sheet, col_num, order | Ordena por coluna |
| `get_filtered_rows()` | sheet | Retorna linhas visíveis |
| `find_row()` | sheet, col, value, data | Encontra primeira linha com valor |

### Tabela Excel

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `list_tables()` | sheet | Lista tabelas |
| `create_table()` | sheet, range, name | Cria tabela |
| `get_table_header_row()` | sheet, table_name | Retorna cabeçalho |
| `read_table()` | sheet, table_name, columns | Lê dados como lista |
| `read_filtered_table()` | sheet, table_name | Lê apenas visíveis |
| `set_table_cell()` | sheet, table_name, row, col, value | Define valor |
| `get_table_cell()` | sheet, table_name, row, col | Retorna valor |

### Copiar/Colar

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `copy()` | — | Copia seleção |
| `cut()` | — | Move seleção |
| `paste()` | — | Cola seleção |
| `paste_special()` | operation | Cola especial |
| `delete()` | — | Deleta seleção |

### Desfazer/Refazer

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `undo()` | — | Desfaz (Ctrl+Z) |
| `redo()` | — | Refaz (Ctrl+Y) |

### Proteção

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `protect_worksheet()` | password | Protege planilha |
| `unprotect_worksheet()` | password | Desprotege planilha |
| `protect_workbook()` | password | Protege pasta |
| `unprotect_workbook()` | password | Desprotege pasta |

### Formatação Condicional

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `add_conditional_format()` | range, rule, format | Adiciona regra |
| `clear_conditional_formats()` | range | Limpa regras |

### Outros

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `set_visible()` | visible | Define visibilidade |
| `get_visible()` | — | Retorna visibilidade |
| `alert()` | message | Exibe alerta |

## Parâmetros Comuns

| Parâmetro | Tipo | Descrição |
|-----------|------|-----------|
| `cell` | str | Referência (ex: "A1") |
| `range` | str | Intervalo (ex: "A1:C10") |
| `sheet_name` | str | Nome da aba |
| `col_num` | int | Número da coluna |
| `row_num` | int | Número da linha |
| `value` | any | Valor a ser escrito |
| `order` | str | "asc" ou "desc" |
| `data` | list | Lista de dados para cache |

## Exemplos Práticos

### Processar planilha grande (Toyota)

```python
from rpaflow.excel_com import ExcelCom

xl = ExcelCom(visible=False)
xl.open(r"C:\Users\sergio.ssilva\Desktop\Dívida Montadora Toyota 30.06.xlsx")

# Ler tabela como lista de dicts
registros = xl.read_filtered_table("pgtos", "Tabela1")
print(f"Registros: {len(registors)}")

# Filtrar por valores
xl.filter_column_values("Vendas", 1, ["PCD"])
visiveis = xl.get_filtered_rows("Vendas")
print(f"Filtrados: {len(visiveis)}")

# Ordenar
xl.sort_column("Vendas", 3, order="desc")

xl.save()
xl.close()
xl.quit()
```

### Usar cache para performance

```python
from rpaflow.excel_com import ExcelCom

xl = ExcelCom(visible=True)
xl.open("C:/dados/producao.xlsx")

# Ler dados uma vez
data = xl.read_table("Pedidos", "Tabela1")

# Reutilizar em múltiplas buscas (30x mais rápido)
for nf in notas_fiscais:
    row = xl.find_row("Pedidos", 5, nf, data=data)
    if row:
        print(f"NF {nf} encontrada na linha {row}")

xl.quit()
```

### Ler tabela com filtros

```python
from rpaflow.excel_com import ExcelCom

xl = ExcelCom(visible=False)
xl.open("C:/dados/vendas.xlsx")

# Aplicar filtro
xl.filter_column_values("Vendas", 1, ["PCD", "RE"])

# Ler apenas linhas visíveis
dados = xl.read_filtered_table("Vendas", "Vendas")
print(f"Registros filtrados: {len(dados)}")

for registro in dados:
    print(f"NF: {registro['NF']}, Valor: {registro['Valor']}")

xl.quit()
```

---

# Browser

Automação de navegador com suporte a múltiplos frames. Inspirado no Browser do Robot Framework.

## Instalar

```bash
pip install rpaflow[browser]
```

## Como Usar (Passo a Passo)

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

## Métodos

### Gerenciamento

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `start()` | url, type, headless | Inicia navegador |
| `close()` | — | Fecha navegador |
| `go_to()` | url | Navega para URL |
| `get_url()` | — | URL atual |
| `get_title()` | — | Título da página |

### Interação

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `click()` | selector, recursive | Clica no elemento |
| `double_click()` | selector | Clique duplo |
| `right_click()` | selector | Clique direito |
| `type_text()` | selector, text | Digita texto |
| `clear_text()` | selector | Limpa campo |
| `select_option()` | selector, value | Seleciona opção |
| `check()` | selector | Marca checkbox |
| `uncheck()` | selector | Desmarca checkbox |
| `hover()` | selector | Passa mouse por cima |
| `press_key()` | key | Pressiona tecla |
| `upload_file()` | selector, filepath | Upload de arquivo |

### Busca

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `find_element()` | selector | Encontra elemento |
| `find_all()` | selector | Encontra todos |
| `is_visible()` | selector, timeout | Verifica visibilidade |
| `is_enabled()` | selector | Verifica se habilitado |
| `get_text()` | selector | Retorna texto |
| `get_attribute()` | selector, attr | Retorna atributo |

### Busca Recursiva

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `find_element_in_frames()` | selector | Encontra em frames |
| `click_in_frames()` | selector | Clica em frames |
| `select_in_frames()` | selector, value | Seleciona em frames |
| `find_all_in_frames()` | selector | Encontra todos em frames |

### Execução

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `execute_js()` | script | Executa JavaScript |
| `wait()` | seconds | Espera fixa |
| `wait_for_element()` | selector, timeout | Espera elemento |

### Navegação

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `back()` | — | Volta |
| `forward()` | — | Avança |
| `reload()` | — | Recarrega |
| `maximize()` | — | Maximiza |

### Imagem

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `screenshot()` | filepath | Tela do elemento |
| `full_screenshot()` | filepath | Tela cheia |

### Propriedade

| Propriedade | Descrição |
|-------------|-----------|
| `page` | Retorna objeto Page do Playwright |

## Parâmetros Comuns

| Parâmetro | Padrão | Descrição |
|-----------|--------|-----------|
| `selector` | — | CSS selector |
| `recursive` | False | Busca recursiva |
| `timeout` | 30000 | Timeout em ms |
| `type` | playwright | Motor do navegador |
| `headless` | False | Modo headless |

## Exemplos Práticos

### Formulário com iframe

```python
from rpaflow.browser import Browser

browser = Browser()
browser.start("https://site.com/form", type="playwright")

# Campo normal
browser.type_text("#nome", "João")

# Campo dentro de iframe (recursive)
browser.type_text("#campo_frame", "valor", recursive=True)

# Clicar botão dentro de iframe
browser.click("#enviar", recursive=True)

browser.close()
```

### Login

```python
from rpaflow.browser import Browser

browser = Browser()
browser.start("https://site.com/login", type="playwright")

browser.type_text("#email", "user@email.com")
browser.type_text("#senha", "123456")
browser.click("#entrar")

# Esperar carregar
browser.wait(3)

# Verificar login
if browser.is_visible("#dashboard"):
    print("Login realizado")

browser.close()
```

### Acessar Page do Playwright

```python
from rpaflow.browser import Browser

browser = Browser()
browser.start("https://site.com", type="playwright")

# Acessar page diretamente
page = browser.page
page.evaluate("alert('Olá!')")
page.screenshot(path="tela.png")

browser.close()
```

---

# Desktop

Automação de aplicativos desktop via reconhecimento de imagem. Inspirado no Library do Robot Framework.

## Instalar

```bash
pip install rpaflow[desktop]
```

## Como Usar

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

## Métodos

### Localização

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `find_image()` | image, confidence | Localiza imagem |
| `find_all_images()` | image, confidence | Localiza todas |
| `wait_for_image()` | image, timeout | Espera imagem |

### Interação

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `click_image()` | image, confidence | Clica na imagem |
| `double_click_image()` | image | Clique duplo |
| `right_click_image()` | image | Clique direito |
| `drag_and_drop()` | source, target | Arrasta |

### Texto

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `type_text()` | text | Digita texto |
| `press_key()` | key | Pressiona tecla |
| `get_clipboard()` | — | Retorna clipboard |
| `set_clipboard()` | text | Define clipboard |

### Janela

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get_active_window()` | — | Janela ativa |
| `set_active_window()` | title | Ativa janela |
| `maximize_window()` | title | Maximiza |
| `minimize_window()` | title | Minimiza |

### Espera

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `wait()` | seconds | Espera fixa |
| `wait_until_image()` | image, timeout | Espera até achar imagem |

## Parâmetros Comuns

| Parâmetro | Padrão | Descrição |
|-----------|--------|-----------|
| `image` | — | Caminho da imagem |
| `confidence` | 0.8 | Confiança (0-1) |
| `timeout` | 10 | Timeout em segundos |

---

# Log

Logging profissional para Python, inspirado no Serilog do C#.

## Instalar

```bash
pip install rpaflow[log]
```

## Como Usar

```python
from rpaflow.log import Log

log = Log(path="C:/logs/meu_bot.log", level="DEBUG", json=True)

log = log.bind(bot="vendas", user="admin")
log.info("Iniciando robô")
log.error("Falha na conexão", host="localhost")
log.success("Robô finalizado")
```

## Métodos

### Criação

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `Log()` | path, level, json | Cria instância |

### Métodos de Log

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `debug()` | message, **kwargs | Log de debug |
| `info()` | message, **kwargs | Log informativo |
| `warning()` | message, **kwargs | Aviso |
| `error()` | message, **kwargs | Erro |
| `critical()` | message, **kwargs | Crítico |
| `success()` | message, **kwargs | Sucesso |

### Bind

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `bind()` | **kwargs | Adiciona contexto |

## Parâmetros Comuns

| Parâmetro | Padrão | Descrição |
|-----------|--------|-----------|
| `path` | — | Caminho do arquivo de log |
| `level` | INFO | Nível mínimo |
| `json` | False | Formato JSON |

## Exemplo Prático

### Log com contexto

```python
from rpaflow.log import Log

log = Log(path="C:/logs/vendas.log", level="DEBUG")
log = log.bind(bot="vendas", user="admin")

log.info("Iniciando processamento")
log.debug("Arquivo carregado", arquivo="vendas.xlsx")
log.info("Processados 150 pedidos")
log.success("Finalizado em 5min")
```

### Log de erros

```python
from rpaflow.log import Log

log = Log(path="C:/logs/erros.log", level="ERROR")

try:
    # código
    pass
except Exception as e:
    log.error("Falha no processamento", erro=str(e))
```

---

# INI

Leitura e escrita de arquivos .ini. Inspirado no `System.IO.File` do C#.

## Instalar

```bash
pip install rpaflow[ini]
```

## Como Usar

```python
from rpaflow.ini import Ini

ini = Ini("C:/config/robô.ini")

host = ini.get("database", "host")
port = ini.get_int("database", "port")
debug = ini.get_bool("app", "debug")

ini.set("database", "host", "192.168.1.100")
ini.save()
```

## Métodos

### Criação

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `Ini()` | filepath | Cria instância |

### Leitura

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get()` | section, key | Lê valor como string |
| `get_int()` | section, key | Lê como inteiro |
| `get_float()` | section, key | Lê como float |
| `get_bool()` | section, key | Lê como booleano |
| `get_list()` | section, key | Lê como lista |

### Escrita

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `set()` | section, key, value | Define valor |
| `remove()` | section, key | Remove chave |
| `remove_section()` | section | Remove seção |
| `save()` | — | Salva arquivo |

### Seções

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `sections()` | — | Lista seções |
| `has_section()` | section | Verifica se existe |
| `has_option()` | section, key | Verifica se chave existe |

## Parâmetros Comuns

| Parâmetro | Padrão | Descrição |
|-----------|--------|-----------|
| `filepath` | — | Caminho do arquivo .ini |
| `section` | — | Nome da seção |
| `key` | — | Nome da chave |
| `value` | — | Valor a ser gravado |

## Formato do Arquivo .ini

```ini
[database]
host=localhost
port=5432
user=admin
password=123456
database=vendas

[app]
debug=true
log_level=INFO
timeout=30

[smtp]
host=smtp.gmail.com
port=587
user=meu@email.com
password=minha_senha
```

## Exemplos Práticos

### Ler configurações

```python
from rpaflow.ini import Ini

ini = Ini("C:/config/robô.ini")

# Ler valores
host = ini.get("database", "host")
port = ini.get_int("database", "port")
debug = ini.get_bool("app", "debug")

print(f"Host: {host}, Port: {port}, Debug: {debug}")
```

### Atualizar configurações

```python
from rpaflow.ini import Ini

ini = Ini("C:/config/robô.ini")

# Atualizar
ini.set("database", "host", "192.168.1.100")
ini.set("app", "debug", "false")
ini.save()

print("Configuração atualizada!")
```

### Criar nova configuração

```python
from rpaflow.ini import Ini

ini = Ini("C:/config/novo.ini")

# Definir valores
ini.set("database", "host", "localhost")
ini.set("database", "port", "5432")
ini.set("database", "user", "admin")
ini.set("database", "password", "123456")
ini.set("database", "database", "vendas")

ini.set("app", "debug", "true")
ini.set("app", "log_level", "INFO")

ini.save()

print("Configuração criada!")
```

### Iterar seções

```python
from rpaflow.ini import Ini

ini = Ini("C:/config/robô.ini")

for section in ini.sections():
    print(f"\n[{section}]")
    for key, value in ini.items(section):
        print(f"  {key} = {value}")
```

### Verificar se chave existe

```python
from rpaflow.ini import Ini

ini = Ini("C:/config/robô.ini")

if ini.has_option("database", "host"):
    print("Host configurado")

if ini.has_section("smtp"):
    print("SMTP configurado")
```

### Criar configuração a partir de dict

```python
from rpaflow.ini import Ini

config = {
    "database": {
        "host": "localhost",
        "port": "5432",
        "user": "admin",
        "password": "123456",
        "database": "vendas"
    },
    "app": {
        "debug": "true",
        "log_level": "INFO",
        "timeout": "30"
    }
}

ini = Ini("C:/config/novo.ini")

for section, options in config.items():
    for key, value in options.items():
        ini.set(section, key, value)

ini.save()

print("Configuração criada a partir de dict!")
```

### Backup de configuração

```python
from rpaflow.ini import Ini
from datetime import datetime

ini = Ini("C:/config/robô.ini")

# Criar backup com data
data = datetime.now().strftime("%Y%m%d_%H%M%S")
backup = f"C:/config/robô_backup_{data}.ini"

import shutil
shutil.copy(ini.filepath, backup)

print(f"Backup criado: {backup}")
```

### Configuração com valores padrão

```python
from rpaflow.ini import Ini

ini = Ini("C:/config/robô.ini")

# Valores padrão caso não existam
defaults = {
    "database": {"host": "localhost", "port": "5432"},
    "app": {"debug": "false", "timeout": "30"}
}

for section, options in defaults.items():
    for key, value in options.items():
        if not ini.has_option(section, key):
            ini.set(section, key, value)

ini.save()
print("Valores padrão aplicados!")
```

### Ler múltiplos arquivos INI

```python
from rpaflow.ini import Ini
import glob

# Ler todos os .ini de uma pasta
arquivos = glob.glob("C:/config/*.ini")

for arquivo in arquivos:
    ini = Ini(arquivo)
    print(f"\nArquivo: {arquivo}")
    for section in ini.sections():
        print(f"  [{section}]")
        for key, value in ini.items(section):
            print(f"    {key} = {value}")
```

### Configuração para múltiplos robôs

```python
from rpaflow.ini import Ini

# Criar config para cada robô
robos = {
    "vendas": {"host": "localhost", "database": "vendas"},
    "estoque": {"host": "192.168.1.100", "database": "estoque"},
    "financeiro": {"host": "10.0.0.1", "database": "financeiro"}
}

for nome, config in robos.items():
    ini = Ini(f"C:/config/{nome}.ini")
    
    for key, value in config.items():
        ini.set("database", key, value)
    
    ini.set("app", "debug", "true")
    ini.save()
    
    print(f"Configuração criada: {nome}")
```

### Merge de configurações

```python
from rpaflow.ini import Ini

# Carregar config base
base = Ini("C:/config/base.ini")

# Carregar config específica
especifica = Ini("C:/config/vendas.ini")

# Merge (especifica sobrescreve base)
for section in especifica.sections():
    for key, value in especifica.items(section):
        base.set(section, key, value)

base.save()

print("Configurações mergeadas!")
```

### Verificar integridade

```python
from rpaflow.ini import Ini

ini = Ini("C:/config/robô.ini")

# Verificar se todas as seções obrigatórias existem
obrigatorias = ["database", "app", "smtp"]
for section in obrigatorias:
    if not ini.has_section(section):
        print(f"Seção obrigatória ausente: {section}")

# Verificar se chaves obrigatórias existem
chaves_obrigatorias = {
    "database": ["host", "port", "user"],
    "app": ["debug", "log_level"]
}

for section, keys in chaves_obrigatorias.items():
    for key in keys:
        if not ini.has_option(section, key):
            print(f"Chave obrigatória ausente: [{section}] {key}")

print("Verificação concluída!")
```

---

# JSON

Leitura e escrita de JSON com dot notation. Inspirado no Newtonsoft.Json do C#.

## Instalar

```bash
pip install rpaflow[json]
```

## Como Usar

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
```

## Métodos

### Criação

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `Json()` | — | Cria instância |

### Leitura

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `load()` | filepath | Lê arquivo JSON |
| `loads()` | text | Lê string JSON |

### Escrita

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `save()` | filepath, data, indent | Salva arquivo JSON |
| `dumps()` | data, indent | Serializa para string |

### Utilitários

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get_keys()` | data | Lista de chaves |
| `get_values()` | data | Lista de valores |
| `get_items()` | data | Lista de tuplas (chave, valor) |
| `has_key()` | data, key | Verifica se chave existe |
| `get_value()` | data, key, default | Valor com fallback |
| `to_dict()` | data | Converte para dict |
| `to_list()` | data | Converte para list |

## Dot Notation

### Acessar valores

```python
data = json.load("config.json")

# Com dot notation (igual Newtonsoft.Json)
data.user.name                       # "Maria"
data.user.address.city               # "SP"
data.database.host                   # "localhost"
data.bots[0].name                    # "invoice-bot"
data.bots[0].enabled                 # True

# Com colchetes (também funciona)
data["user"]["name"]                 # "Maria"
data["bots"][0]["name"]              # "invoice-bot"
```

### Setar valores

```python
data.user.name = "João"
data.user.phone = "+55 11 99999-0000"
data.user.address.country = "Brazil"  # Cria automaticamente
```

### Iterar arrays

```python
# Array de objetos
for bot in data.bots:
    print(f"Bot: {bot.name}")
    print(f"  Enabled: {bot.enabled}")
    print(f"  Retry: {bot.retry}")

# Dicionário
for empresa, config in data.empresas.items():
    print(f"{empresa}: {config.host}:{config.port}")
```

## Exemplos Práticos

### Iterar empresas e conectar ao banco

```python
from rpaflow.json import Json
from rpaflow.sql import SQL
from rpaflow.log import Log

json = Json()
log = Log(path="C:/logs/robô.log", level="DEBUG")

data = json.load("C:/config/empresas.json")

for empresa, config in data.empresas.items():
    log.info(f"Processando empresa: {empresa}")
    
    sql = SQL(
        type="mysql",
        host=config.host,
        user=config.user,
        password=config.password,
        database=config.database
    )
    sql.connect()
    
    pedidos = sql.select("pedidos", where={"status": "pendente"})
    log.info(f"Encontrados {len(pedidos)} pedidos")
    
    sql.disconnect()
```

### Iterar bots e executar

```python
from rpaflow.json import Json
from rpaflow.log import Log

json = Json()
log = Log(path="C:/logs/bots.log", level="DEBUG")

data = json.load("C:/config/bots.json")

for bot in data.bots:
    if bot.enabled:
        log.info(f"Executando bot: {bot.name}")
        # ... executar bot
        log.success(f"Bot {bot.name} finalizado")
    else:
        log.warning(f"Bot {bot.name} desabilitado")
```

### Carregar configuração de robô

```python
from rpaflow.json import Json
from rpaflow.sql import SQL
from rpaflow.log import Log
from rpaflow.excel_com import ExcelCom

json = Json()
data = json.load("C:/config/robô.json")

# Configurações do banco
db = data.database
sql = SQL(
    type=db.type,
    host=db.host,
    user=db.user,
    password=db.password,
    database=db.database
)

# Configurações de log
log_config = data.logging
log = Log(
    path=log_config.file,
    level=log_config.level,
    json=log_config.json
)

# Configurações do Excel
excel_config = data.excel
xl = ExcelCom(visible=excel_config.visible)
xl.open(excel_config.path)

log.info("Robô iniciado")
sql.connect()
# ... processar
log.success("Robô finalizado")
```

### Salvar dados processados

```python
from rpaflow.json import Json

json = Json()

# Criar dados
resultado = {
    "processados": 150,
    "erros": 3,
    "tempo_execucao": "00:05:30",
    "detalhes": [
        {"pedido": 1001, "status": "ok"},
        {"pedido": 1002, "status": "ok"},
        {"pedido": 1003, "status": "erro", "motivo": "timeout"}
    ]
}

# Salvar
json.save("C:/logs/resultado.json", resultado, indent=4)
print("Resultado salvo")
```

### Atualizar configuração existente

```python
from rpaflow.json import Json

json = Json()
data = json.load("C:/config/app.json")

# Atualizar valores
data.database.host = "192.168.1.100"
data.database.port = 5432

# Adicionar novo campo
data.database.timeout = 30

# Salvar
json.save("C:/config/app.json", data)
print("Configuração atualizada")
```

## Comparação com C# Newtonsoft.Json

| C# Newtonsoft.Json | Python rpaflow.json |
|--------------------|---------------------|
| `JsonConvert.DeserializeObject<T>(json)` | `json.loads(text)` |
| `JsonConvert.SerializeObject(obj)` | `json.dumps(data)` |
| `File.ReadAllText("config.json")` | `json.load("config.json")` |
| `File.WriteAllText("config.json", json)` | `json.save("config.json", data)` |
| `obj.User.Name` | `data.user.name` |
| `obj.Bots[0].Name` | `data.bots[0].name` |
| `foreach (var bot in obj.Bots)` | `for bot in data.bots:` |
| `foreach (var kvp in obj.Empresas)` | `for empresa, config in data.empresas.items():` |

---

# Files

Operações com arquivos e diretórios. Inspirado nos métodos estáticos de `System.IO.File` e `System.IO.Path` do C#.

## Instalar

```bash
pip install rpaflow[files]
```

## Como Usar

```python
from rpaflow.files import Files

files = Files()

# Leitura/Escrita
conteudo = files.read_text("arquivo.txt")
files.write_text("arquivo.txt", "conteúdo")
files.append_text("arquivo.txt", "novo conteúdo")

# Linhas
linhas = files.read_lines("arquivo.txt")
files.write_lines("arquivo.txt", ["linha1", "linha2"])

# Operações
files.exists("arquivo.txt")
files.copy("origem.txt", "destino.txt")
files.move("antigo.txt", "novo.txt")
files.delete("lixo.txt")

# Path
files.get_filename("C:/pasta/arquivo.txt")
files.get_extension("arquivo.txt")
files.combine("C:/pasta", "arquivo.txt")
```

## Métodos

### Leitura/Escrita Texto

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `read_text()` | filepath, encoding | Lê conteúdo como texto |
| `write_text()` | filepath, content, encoding | Escreve texto no arquivo |
| `append_text()` | filepath, content, encoding | Adiciona texto ao final |
| `read_lines()` | filepath, encoding | Lê linhas como lista |
| `write_lines()` | filepath, lines, encoding | Escreve lista de linhas |
| `append_lines()` | filepath, lines, encoding | Adiciona linhas ao final |

### Leitura/Escrita Bytes

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `read_bytes()` | filepath | Lê conteúdo como bytes |
| `write_bytes()` | filepath, data | Escreve bytes no arquivo |
| `append_bytes()` | filepath, data | Adiciona bytes ao final |

### Operações de Arquivo

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `exists()` | filepath | Verifica se arquivo existe |
| `copy()` | src, dst, overwrite | Copia arquivo |
| `move()` | src, dst, overwrite | Move arquivo |
| `delete()` | filepath | Deleta arquivo |
| `replace()` | source, destination, backup | Substitui arquivo |

### Informações do Arquivo

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get_creation_time()` | filepath | Data de criação |
| `get_last_write_time()` | filepath | Data da última escrita |
| `get_last_access_time()` | filepath | Data do último acesso |
| `get_attributes()` | filepath | Atributos do arquivo |
| `set_attributes()` | filepath, readonly, hidden | Define atributos |
| `create_symlink()` | link, target | Cria link simbólico |

### Path

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get_filename()` | filepath | Nome do arquivo com extensão |
| `get_filename_without_ext()` | filepath | Nome sem extensão |
| `get_extension()` | filepath | Extensão do arquivo |
| `get_directory()` | filepath | Diretório pai |
| `combine()` | *paths | Junta caminhos |
| `join()` | *paths | Junta caminhos (os.path.join) |
| `get_full_path()` | filepath | Caminho absoluto |
| `get_relative_path()` | base, filepath | Caminho relativo |
| `is_rooted()` | filepath | Verifica se é absoluto |
| `is_fully_qualified()` | filepath | Verifica se é completo |
| `has_extension()` | filepath | Verifica se tem extensão |
| `change_extension()` | filepath, new_ext | Troca extensão |
| `trim_separator()` | filepath | Remove separador final |
| `ends_with_separator()` | filepath | Verifica se termina com separador |
| `get_temp_path()` | — | Diretório temporário |
| `get_temp_file()` | — | Arquivo temporário |
| `get_random_filename()` | — | Nome aleatório |
| `get_invalid_filename_chars()` | — | Caracteres inválidos no nome |
| `get_invalid_path_chars()` | — | Caracteres inválidos no caminho |
| `get_path_root()` | filepath | Raiz do caminho |
| `get_base_directory()` | — | Diretório base (AppDomain.BaseDirectory) |

### Directory

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `dir_create()` | path | Cria diretório |
| `dir_delete()` | path, recursive | Deleta diretório |
| `dir_exists()` | path | Verifica se diretório existe |
| `dir_move()` | src, dst | Move diretório |
| `dir_get_files()` | path, pattern | Lista arquivos |
| `dir_get_dirs()` | path | Lista subdiretórios |
| `dir_get_all_files()` | path, pattern | Lista todos recursivamente |

### File Info

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get_size()` | filepath | Tamanho em bytes |
| `is_readonly()` | filepath | Verifica se é somente leitura |
| `set_readonly()` | filepath, readonly | Define somente leitura |

## Exemplos Práticos

### Ler e processar arquivo CSV

```python
from rpaflow.files import Files

files = Files()

# Ler CSV como linhas
linhas = files.read_lines("dados.csv", encoding="latin-1")

for linha in linhas:
    campos = linha.split(";")
    print(f"Nome: {campos[0]}, Idade: {campos[1]}")
```

### Backup de arquivo

```python
from rpaflow.files import Files
from datetime import datetime

files = Files()

# Criar backup com data
data = datetime.now().strftime("%Y%m%d_%H%M%S")
backup = f"backup_{data}.txt"
files.copy("arquivo.txt", backup)

print(f"Backup criado: {backup}")
```

### Processar todos os arquivos de uma pasta

```python
from rpaflow.files import Files

files = Files()

# Listar todos os .txt recursivamente
arquivos = files.dir_get_all_files("C:/dados", "*.txt")

for arquivo in arquivos:
    conteudo = files.read_text(arquivo)
    print(f"{arquivo}: {len(conteudo)} caracteres")
```

### Verificar e criar diretório

```python
from rpaflow.files import Files

files = Files()

if not files.dir_exists("C:/logs"):
    files.dir_create("C:/logs")
    print("Diretório criado")

# Escrever log
files.append_text("C:/logs/app.log", "Nova entrada de log\n")
```

### Trabalhar com caminhos

```python
from rpaflow.files import Files

files = Files()

caminho = "C:/pasta/subpasta/arquivo.txt"

print(files.get_filename(caminho))              # "arquivo.txt"
print(files.get_filename_without_ext(caminho))  # "arquivo"
print(files.get_extension(caminho))             # ".txt"
print(files.get_directory(caminho))             # "C:/pasta/subpasta"
print(files.get_path_root(caminho))             # "C:\"
print(files.is_rooted(caminho))                 # True

# Combinar caminhos
novo = files.combine("C:/pasta", "subpasta", "arquivo.txt")
print(novo)  # "C:/pasta/subpasta/arquivo.txt"

# Trocar extensão
log = files.change_extension(caminho, ".log")
print(log)  # "C:/pasta/subpasta/arquivo.log"
```

---

# API

Requisições HTTP/REST.

## Instalar

```bash
pip install rpaflow[api]
```

## Como Usar

```python
from rpaflow.api import API

api = API()

# GET
response = api.get("https://api.example.com/users")

# POST
api.post("https://api.example.com/users", json={"nome": "João"})

# PUT
api.put("https://api.example.com/users/1", json={"nome": "João Silva"})

# DELETE
api.delete("https://api.example.com/users/1")
```

## Métodos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get()` | url, headers (opcional) | Requisição GET |
| `post()` | url, json, headers (opcional) | Requisição POST |
| `put()` | url, json, headers (opcional) | Requisição PUT |
| `delete()` | url, headers (opcional) | Requisição DELETE |

---

# Email

Envio de emails via SMTP.

## Instalar

```bash
pip install rpaflow[email]
```

## Como Usar

```python
from rpaflow.email import Email

email = Email(smtp_host="smtp.gmail.com", smtp_port=587, user="meu@email.com", password="senha")

# Conectar
email.connect()

# Enviar
email.send(
    to="dest@email.com",
    subject="Assunto",
    body="Mensagem do email",
    attachments=["arquivo.pdf"]
)

# Desconectar
email.disconnect()
```

## Métodos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `connect()` | — | Conecta ao SMTP |
| `send()` | to, subject, body, attachments | Envia email |
| `disconnect()` | — | Desconecta |

---

## Licença

MIT License - veja [LICENSE](LICENSE) para detalhes.
