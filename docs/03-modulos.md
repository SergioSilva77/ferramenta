# 03 - Módulos

## SQL

Suporta MySQL, PostgreSQL, SQL Server e SQLite.

```python
from rpaflow.sql import SQL

# MySQL
db = SQL(type="mysql", host="localhost", user="root", password="123", database="vendas")

# PostgreSQL
db = SQL(type="postgresql", host="localhost", user="postgres", password="123", database="vendas")

# SQL Server
db = SQL(type="sqlserver", host="localhost", user="sa", password="123", database="vendas")

# SQLite (sem autenticação)
db = SQL(type="sqlite", database="meubanco.db")

# Conectar
db.connect()

# Inserir
db.insert("clientes", {"nome": "João", "email": "joao@email.com"})

# Selecionar
clientes = db.select("clientes")
clientes_ativos = db.select("clientes", where={"ativo": True})

# Atualizar
db.update("clientes", {"nome": "João Silva"}, where={"id": 1})

# Deletar
db.delete("clientes", where={"id": 1})

# Executar query customizada
resultados = db.execute("SELECT * FROM clientes WHERE id > %s", [10])

# Desconectar
db.disconnect()
```

### Métodos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `connect()` | — | Conecta ao banco |
| `select()` | table, where (dict, opcional) | Retorna registros |
| `insert()` | table, data (dict) | Insere registro |
| `update()` | table, data (dict), where (dict) | Atualiza registros |
| `delete()` | table, where (dict) | Deleta registros |
| `execute()` | query, params (list, opcional) | Executa query customizada |
| `disconnect()` | — | Fecha conexão |

### Instalar Dependências

```bash
pip install rpaflow[sql]              # MySQL
pip install rpaflow[sql-postgresql]   # PostgreSQL
pip install rpaflow[sql-sqlserver]    # SQL Server
pip install rpaflow[sql-all]          # Todos
```

### Suporte por Banco

| Banco | Tipo | Biblioteca | Porta Padrão |
|-------|------|------------|--------------|
| MySQL | `mysql` | pymysql | 3306 |
| PostgreSQL | `postgresql` | psycopg2-binary | 5432 |
| SQL Server | `sqlserver` | pyodbc | 1433 |
| SQLite | `sqlite` | sqlite3 (built-in) | — |

---

## Excel

```python
from rpaflow.excel import Excel

# Abrir planilha
planilha = Excel("dados.xlsx")
planilha.open()

# Ler dados
dados = planilha.read("Planilha1", range="A1:D10")

# Escrever dados
planilha.write("Planilha1", range="A1", values=[
    ["Nome", "Idade", "Cidade"],
    ["Ana", 25, "São Paulo"],
    ["João", 30, "Rio"]
])

# Salvar
planilha.save("saida.xlsx")

# Fechar
planilha.close()
```

### Métodos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `open()` | — | Abre planilha |
| `read()` | sheet, range (opcional) | Lê dados |
| `write()` | sheet, range, values | Escreve dados |
| `save()` | filepath | Salva arquivo |
| `close()` | — | Fecha planilha |

---

## Browser

```python
from rpaflow.browser import Browser

# Iniciar navegador
browser = Browser()
browser.start(url="https://site.com", type="playwright")

# Clicar
browser.click(selector="#botao")

# Digitar
browser.type(selector="#input", text="Olá")

# Ler texto
texto = browser.get_text(selector="h1")

# Screenshot
browser.screenshot("captura.png")

# Fechar
browser.close()
```

### Métodos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `start()` | url, type (playwright/selenium) | Abre navegador |
| `click()` | selector | Clica em elemento |
| `type()` | selector, text | Digita em campo |
| `get_text()` | selector | Retorna texto |
| `screenshot()` | filepath | Captura tela |
| `close()` | — | Fecha navegador |

---

## Files

```python
from rpaflow.files import Files

files = Files()

# Ler
conteudo = files.read("arquivo.txt")

# Escrever
files.write("saida.txt", "conteúdo do arquivo")

# Copiar
files.copy("origem.txt", "destino.txt")

# Mover
files.move("antigo.txt", "pasta/novo.txt")

# Deletar
files.delete("lixo.txt")
```

### Métodos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `read()` | filepath | Lê conteúdo |
| `write()` | filepath, content | Escreve conteúdo |
| `copy()` | src, dst | Copia arquivo |
| `move()` | src, dst | Move arquivo |
| `delete()` | filepath | Deleta arquivo |

---

## API

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

### Métodos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get()` | url, headers (opcional) | Requisição GET |
| `post()` | url, json, headers (opcional) | Requisição POST |
| `put()` | url, json, headers (opcional) | Requisição PUT |
| `delete()` | url, headers (opcional) | Requisição DELETE |

---

## Email

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

### Métodos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `connect()` | — | Conecta ao SMTP |
| `send()` | to, subject, body, attachments | Envia email |
| `disconnect()` | — | Desconecta |
