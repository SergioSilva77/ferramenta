# SQL

Suporta MySQL, PostgreSQL, SQL Server e SQLite.

## Instalar

```bash
pip install rpaflow[sql]              # MySQL
pip install rpaflow[sql-postgresql]   # PostgreSQL
pip install rpaflow[sql-sqlserver]    # SQL Server
pip install rpaflow[sql-all]          # Todos
```

## Suporte por Banco

| Banco | Tipo | Biblioteca | Porta Padrão |
|-------|------|------------|--------------|
| MySQL | `mysql` | pymysql | 3306 |
| PostgreSQL | `postgresql` | psycopg2-binary | 5432 |
| SQL Server | `sqlserver` | pyodbc | 1433 |
| SQLite | `sqlite` | sqlite3 (built-in) | — |

## Exemplo

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

## Métodos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `connect()` | — | Conecta ao banco |
| `select()` | table, where (dict, opcional) | Retorna registros |
| `insert()` | table, data (dict) | Insere registro |
| `update()` | table, data (dict), where (dict) | Atualiza registros |
| `delete()` | table, where (dict) | Deleta registros |
| `execute()` | query, params (list, opcional) | Executa query customizada |
| `disconnect()` | — | Fecha conexão |
