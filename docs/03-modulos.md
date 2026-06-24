# 03 - Módulos

## SQL

```python
from rpaflow.sql import SQL

# Conectar
db = SQL(host="localhost", user="root", password="123", database="vendas", type="mysql")
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

# Desconectar
db.disconnect()
```

### Métodos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `connect()` | — | Conecta ao banco |
| `insert()` | table, data (dict) | Insere registro |
| `select()` | table, where (dict, opcional) | Retorna registros |
| `update()` | table, data (dict), where (dict) | Atualiza registros |
| `delete()` | table, where (dict) | Deleta registros |
| `disconnect()` | — | Fecha conexão |

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

email = Email()

# Enviar
email.send(
    to="dest@email.com",
    subject="Assunto",
    body="Mensagem do email",
    attachments=["arquivo.pdf"]
)

# Ler
emails = email.read(folder="INBOX", limit=10)
```

### Métodos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `send()` | to, subject, body, attachments | Envia email |
| `read()` | folder, limit | Lê emails |
