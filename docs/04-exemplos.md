# 04 - Exemplos

## Exemplo 1: SQL + Excel

Exportar dados do banco para planilha.

```python
from rpaflow import sql, excel

# Conectar ao banco
sql.connect(host="localhost", user="root", password="123", database="vendas", type="mysql")

# Buscar dados
clientes = sql.select("clientes")
sql.disconnect()

# Criar planilha
excel.open("clientes.xlsx")
excel.write("Planilha1", range="A1", values=[
    ["ID", "Nome", "Email", "Ativo"]
])

for i, cliente in enumerate(clientes, start=2):
    excel.write("Planilha1", range=f"A{i}", values=[[cliente[0], cliente[1], cliente[2], cliente[3]]])

excel.save("clientes.xlsx")
excel.close()

print("Planilha gerada com sucesso!")
```

---

## Exemplo 2: Browser + SQL

Fazer login em um site e salvar dados no banco.

```python
from rpaflow import browser, sql

# Iniciar navegador
browser.start(url="https://site.com/login", type="playwright")

# Fazer login
browser.type(selector="#user", text="admin")
browser.type(selector="#pass", text="123456")
browser.click(selector="#btn-login")

# Extrair dados
titulo = browser.get_text(selector="h1")
browser.screenshot("apos_login.png")

# Salvar no banco
sql.connect(host="localhost", user="root", password="123", database="logs", type="mysql")
sql.insert("acessos", {"usuario": "admin", "pagina": titulo})
sql.disconnect()

browser.close()
print("Login e registro concluídos!")
```

---

## Exemplo 3: API + Files

Buscar dados de uma API e salvar em arquivo.

```python
from rpaflow import api, files

# Buscar dados da API
response = api.get("https://api.example.com/users")
usuarios = response.json()

# Salvar em arquivo
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
from rpaflow import email, excel

# Gerar planilha
excel.open("relatorio.xlsx")
excel.write("Relatório", range="A1", values=[
    ["Produto", "Quantidade", "Valor"],
    ["Notebook", 10, 4500.00],
    ["Mouse", 50, 25.00],
    ["Teclado", 30, 75.00]
])
excel.save("relatorio.xlsx")
excel.close()

# Enviar por email
email.send(
    to="gerente@empresa.com",
    subject="Relatório de Estoque",
    body="Segue o relatório em anexo.",
    attachments=["relatorio.xlsx"]
)

print("Email enviado com sucesso!")
```

---

## Exemplo 5: Files + Browser

Baixar arquivos de um site.

```python
from rpaflow import browser, files

# Navegar até a página
browser.start(url="https://site.com/downloads", type="playwright")

# Encontrar links
links = browser.get_all(selector="a.download-link")

for link in links:
    href = link.get_attribute("href")
    filename = href.split("/")[-1]
    print(f"Baixando: {filename}")
    # Download via API ou outro método

browser.close()
print("Downloads concluídos!")
```

---

## Exemplo 6: Try/Catch com Vários Módulos

Tratamento de erros em automação complexa.

```python
from rpaflow import sql, browser, excel

try:
    # Conectar ao banco
    sql.connect(host="localhost", user="root", password="123", database="vendas", type="mysql")

    # Buscar pedidos
    pedidos = sql.select("pedidos", where={"status": "pendente"})
    print(f"{len(pedidos)} pedidos encontrados")

    # Processar cada pedido
    for pedido in pedidos:
        browser.start(url=f"https://sistema.com/pedido/{pedido[0]}", type="playwright")
        browser.click(selector="#aprovar")
        browser.screenshot(f"pedido_{pedido[0]}.png")
        browser.close()

        sql.update("pedidos", {"status": "aprovado"}, where={"id": pedido[0]})

    # Gerar relatório
    excel.open("relatorio_aprovados.xlsx")
    excel.write("Pedidos", range="A1", values=pedidos)
    excel.save("relatorio_aprovados.xlsx")
    excel.close()

    print("Automação concluída com sucesso!")

except Exception as e:
    print(f"Erro: {e}")

finally:
    try:
        sql.disconnect()
        browser.close()
    except:
        pass
```
