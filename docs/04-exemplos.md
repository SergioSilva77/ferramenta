# 04 - Exemplos

## Exemplo 1: SQL + Excel

Exportar dados do banco para planilha.

```python
from rpaflow.sql import SQL
from rpaflow.excel import Excel

# Conectar ao banco
db = SQL(host="localhost", user="root", password="123", database="vendas", type="mysql")
db.connect()

# Buscar dados
clientes = db.select("clientes")
db.disconnect()

# Criar planilha
planilha = Excel("clientes.xlsx")
planilha.open()
planilha.write("Planilha1", range="A1", values=[
    ["ID", "Nome", "Email", "Ativo"]
])

for i, cliente in enumerate(clientes, start=2):
    planilha.write("Planilha1", range=f"A{i}", values=[[cliente[0], cliente[1], cliente[2], cliente[3]]])

planilha.save("clientes.xlsx")
planilha.close()

print("Planilha gerada com sucesso!")
```

---

## Exemplo 2: Browser + SQL

Fazer login em um site e salvar dados no banco.

```python
from rpaflow.browser import Browser
from rpaflow.sql import SQL

# Iniciar navegador
browser = Browser()
browser.start(url="https://site.com/login", type="playwright")

# Fazer login
browser.type(selector="#user", text="admin")
browser.type(selector="#pass", text="123456")
browser.click(selector="#btn-login")

# Extrair dados
titulo = browser.get_text(selector="h1")
browser.screenshot("apos_login.png")

# Salvar no banco
db = SQL(host="localhost", user="root", password="123", database="logs", type="mysql")
db.connect()
db.insert("acessos", {"usuario": "admin", "pagina": titulo})
db.disconnect()

browser.close()
print("Login e registro concluídos!")
```

---

## Exemplo 3: API + Files

Buscar dados de uma API e salvar em arquivo.

```python
from rpaflow.api import API
from rpaflow.files import Files

# Buscar dados da API
api = API()
response = api.get("https://api.example.com/users")
usuarios = response.json()

# Salvar em arquivo
files = Files()
conteudo = ""
for user in usuarios:
    conteudo += f"{user['name']} - {user['email']}\n"

files.write("usuarios.txt", conteudo)
print(f"{len(usuarios)} usuários salvos em usuarios.txt")
```

---

## Exemplo 4: Email + Excel

Enviar planilha por email.

```python
from rpaflow.email import Email
from rpaflow.excel import Excel

# Gerar planilha
planilha = Excel("relatorio.xlsx")
planilha.open()
planilha.write("Relatório", range="A1", values=[
    ["Produto", "Quantidade", "Valor"],
    ["Notebook", 10, 4500.00],
    ["Mouse", 50, 25.00],
    ["Teclado", 30, 75.00]
])
planilha.save("relatorio.xlsx")
planilha.close()

# Enviar por email
email = Email()
email.send(
    to="gerente@empresa.com",
    subject="Relatório de Estoque",
    body="Segue o relatório em anexo.",
    attachments=["relatorio.xlsx"]
)

print("Email enviado com sucesso!")
```

---

## Exemplo 5: Automação Completa

```python
from rpaflow.sql import SQL
from rpaflow.browser import Browser
from rpaflow.excel import Excel

try:
    # Conectar ao banco
    db = SQL(host="localhost", user="root", password="123", database="vendas", type="mysql")
    db.connect()

    # Buscar pedidos pendentes
    pedidos = db.select("pedidos", where={"status": "pendente"})
    print(f"{len(pedidos)} pedidos encontrados")

    # Processar cada pedido
    for pedido in pedidos:
        browser = Browser()
        browser.start(url=f"https://sistema.com/pedido/{pedido[0]}", type="playwright")
        browser.click(selector="#aprovar")
        browser.screenshot(f"pedido_{pedido[0]}.png")
        browser.close()

        db.update("pedidos", {"status": "aprovado"}, where={"id": pedido[0]})

    # Gerar relatório
    planilha = Excel("relatorio_aprovados.xlsx")
    planilha.open()
    planilha.write("Pedidos", range="A1", values=pedidos)
    planilha.save("relatorio_aprovados.xlsx")
    planilha.close()

    print("Automação concluída com sucesso!")

except Exception as e:
    print(f"Erro: {e}")

finally:
    try:
        db.disconnect()
    except:
        pass
```
