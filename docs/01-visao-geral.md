# 01 - Visão Geral

## O que é rpaflow

Biblioteca Python modular para automação RPA. Inspiração no BotCity SDK — cada funcionalidade é um módulo importável via `pip install`.

## Objetivo

Eliminar boilerplate em automações RPA. Em vez de escrever 200+ linhas toda vez:

```python
# Sem rpaflow
import pymysql
import openpyxl
from playwright.sync_api import sync_playwright

conn = pymysql.connect(host="localhost", user="root", password="123", database="vendas")
cursor = conn.cursor()
cursor.execute("SELECT * FROM clientes")
dados = cursor.fetchall()
conn.close()

wb = openpyxl.load_workbook("saida.xlsx")
ws = wb.active
for i, row in enumerate(dados, start=1):
    ws.cell(row=i, column=1, value=row[0])
wb.save("saida.xlsx")
```

```python
# Com rpaflow
from rpaflow import sql, excel

sql.connect(host="localhost", user="root", password="123", database="vendas", type="mysql")
dados = sql.select("clientes")
sql.disconnect()

excel.open("saida.xlsx")
excel.write("Planilha1", range="A1", values=dados)
excel.save("saida.xlsx")
```

## Por que rpaflow?

- **Simples**: Uma linha para cada operação
- **Modular**: Instala só o que precisa
- **Extensível**: Adiciona novos módulos conforme necessidade
- **Consistente**: Mesma API para todos os módulos
