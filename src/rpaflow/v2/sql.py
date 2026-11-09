"""Módulo SQL v2 para rpaflow.

Diferenças em relação ao `rpaflow.sql.SQL` (v1), as principais:

- `QueryError` agora é importado (no v1 toda query com erro virava `NameError` e escondia o erro real).
- INSERT/UPDATE/DELETE funcionam em SQL Server e PostgreSQL: o v1 chamava `fetchall()` em tudo e o driver
  respondia "no results to fetch", o que disparava ROLLBACK — ou seja, o dado nunca era gravado.
- `execute()` devolve linhas só quando a query retorna linhas; `rowcount` fica em `last_rowcount`.
- Cursor sempre fechado; `rollback` que falha (conexão caída) não mascara o erro original.
- Nomes de tabela/coluna são validados (o v1 concatenava texto cru no SQL).
- `driver` do SQL Server é parâmetro (o v1 fixava "ODBC Driver 17"); `trusted=True` usa login do Windows.
- `?` funciona em qualquer banco: é convertido para o placeholder do driver (`%s` no MySQL/PostgreSQL).
- `with SQL(...) as db:` conecta e desconecta sozinho; `execute_many`, `fetch_dicts` e `transaction()` novos.
"""

from __future__ import annotations

import re
from contextlib import contextmanager
from typing import Any, Optional

from rpaflow._exceptions import ConnectionError as RPAConnectionError
from rpaflow._exceptions import QueryError

_IDENTIFICADOR = re.compile(r"^[A-Za-z_][A-Za-z0-9_$]*(\.[A-Za-z_][A-Za-z0-9_$]*){0,2}$")


def _ident(nome: str) -> str:
    """Valida nome de tabela/coluna (aceita schema.tabela). Não dá pra parametrizar identificador, então valida."""
    if not isinstance(nome, str) or not _IDENTIFICADOR.match(nome):
        raise ValueError(f"Identificador SQL inválido: {nome!r}")
    return nome


class SQL:
    """Classe para operações com bancos de dados SQL.

    Tipos suportados: mysql, postgresql, sqlserver, sqlite
    """

    _PORTAS = {"mysql": 3306, "postgresql": 5432, "sqlserver": 1433, "sqlite": 0}

    def __init__(
        self,
        host: str = "localhost",
        user: str = "",
        password: str = "",
        database: str = "",
        port: int = 0,
        type: str = "mysql",
        driver: str = "ODBC Driver 17 for SQL Server",
        trusted: bool = False,
        timeout: int = 30,
    ):
        """
        Args:
            driver: driver ODBC do SQL Server ("ODBC Driver 18 for SQL Server", etc.).
            trusted: SQL Server com autenticação do Windows (ignora user/password).
            timeout: segundos para conectar.
        """
        self.host = host
        self.user = user
        self.password = password
        self.database = database
        self.type = type.lower()
        self.driver = driver
        self.trusted = trusted
        self.timeout = timeout
        self.port = port or self._PORTAS.get(self.type, 3306)
        self.last_rowcount = -1
        self._conn = None

    # ========== CONEXÃO ==========

    def connect(self) -> bool:
        """Conecta ao banco de dados."""
        if self._conn is not None:
            return True
        try:
            self._conn = self._get_connection()
            return True
        except Exception as e:
            raise RPAConnectionError(f"Erro ao conectar: {e}") from e

    def _get_connection(self):
        if self.type == "mysql":
            import pymysql
            return pymysql.connect(
                host=self.host, user=self.user, password=self.password,
                database=self.database, port=self.port, connect_timeout=self.timeout,
            )

        if self.type == "postgresql":
            import psycopg2
            return psycopg2.connect(
                host=self.host, user=self.user, password=self.password,
                dbname=self.database, port=self.port, connect_timeout=self.timeout,
            )

        if self.type == "sqlserver":
            import pyodbc
            auth = "Trusted_Connection=yes;" if self.trusted else f"UID={self.user};PWD={self.password};"
            conn_str = (
                f"DRIVER={{{self.driver}}};SERVER={self.host},{self.port};"
                f"DATABASE={self.database};{auth}"
            )
            return pyodbc.connect(conn_str, timeout=self.timeout)

        if self.type == "sqlite":
            import sqlite3
            return sqlite3.connect(self.database, timeout=self.timeout)

        raise ValueError(f"Tipo '{self.type}' não suportado. Use: mysql, postgresql, sqlserver ou sqlite.")

    def disconnect(self) -> None:
        """Desconecta do banco de dados."""
        if self._conn:
            try:
                self._conn.close()
            finally:
                self._conn = None

    def __enter__(self) -> "SQL":
        self.connect()
        return self

    def __exit__(self, *exc) -> bool:
        self.disconnect()
        return False

    # ========== SQL PRONTO ==========

    def _get_placeholder(self) -> str:
        return "%s" if self.type in ("mysql", "postgresql") else "?"

    def select(self, table: str, where: Optional[dict] = None) -> list:
        """Seleciona registros de uma tabela."""
        ph = self._get_placeholder()
        query = f"SELECT * FROM {_ident(table)}"
        params = []
        if where:
            query += " WHERE " + " AND ".join(f"{_ident(k)} = {ph}" for k in where)
            params = list(where.values())
        return self._execute(query, params)

    def insert(self, table: str, data: dict) -> bool:
        """Insere um registro na tabela."""
        if not data:
            raise ValueError("data vazio")
        ph = self._get_placeholder()
        colunas = ", ".join(_ident(k) for k in data)
        placeholders = ", ".join([ph] * len(data))
        self._execute(f"INSERT INTO {_ident(table)} ({colunas}) VALUES ({placeholders})", list(data.values()))
        return True

    def update(self, table: str, data: dict, where: dict) -> bool:
        """Atualiza registros na tabela. `where` é obrigatório (sem ele o UPDATE pegaria a tabela toda)."""
        if not data:
            raise ValueError("data vazio")
        if not where:
            raise ValueError("where vazio: recusado para não atualizar a tabela inteira")
        ph = self._get_placeholder()
        set_clause = ", ".join(f"{_ident(k)} = {ph}" for k in data)
        where_clause = " AND ".join(f"{_ident(k)} = {ph}" for k in where)
        self._execute(
            f"UPDATE {_ident(table)} SET {set_clause} WHERE {where_clause}",
            list(data.values()) + list(where.values()),
        )
        return True

    def delete(self, table: str, where: dict) -> bool:
        """Deleta registros da tabela. `where` é obrigatório."""
        if not where:
            raise ValueError("where vazio: recusado para não apagar a tabela inteira")
        ph = self._get_placeholder()
        where_clause = " AND ".join(f"{_ident(k)} = {ph}" for k in where)
        self._execute(f"DELETE FROM {_ident(table)} WHERE {where_clause}", list(where.values()))
        return True

    # ========== SQL LIVRE ==========

    def execute(self, query: str, params: Optional[list] = None) -> list:
        """Executa uma query. Devolve as linhas (lista vazia se a query não retorna linhas).

        Use `?` como placeholder em qualquer banco. Ex: db.execute("SELECT * FROM t WHERE id = ?", [1])
        """
        return self._execute(query, params)

    def fetch_dicts(self, query: str, params: Optional[list] = None) -> list:
        """Como execute(), mas cada linha vira dict {coluna: valor}."""
        return self._execute(query, params, as_dict=True)

    def execute_many(self, query: str, rows: list) -> int:
        """Executa a mesma instrução para várias linhas numa transação só. Retorna o rowcount."""
        self._garantir_conexao()
        cursor = self._conn.cursor()
        try:
            cursor.executemany(self._traduzir(query), [list(r) for r in rows])
            self._conn.commit()
            self.last_rowcount = cursor.rowcount
            return self.last_rowcount
        except Exception as e:
            self._rollback()
            raise QueryError(f"Erro ao executar query: {e}") from e
        finally:
            self._fechar(cursor)

    @contextmanager
    def transaction(self):
        """Várias operações com UM commit no fim (ou rollback se algo falhar).

            with db.transaction():
                db.insert(...)
                db.update(...)
        """
        self._garantir_conexao()
        self._em_transacao = True
        try:
            yield self
            self._conn.commit()
        except Exception:
            self._rollback()
            raise
        finally:
            self._em_transacao = False

    _em_transacao = False

    # ========== INTERNOS ==========

    def _garantir_conexao(self) -> None:
        if not self._conn:
            raise RPAConnectionError("Não conectado ao banco. Chame connect() primeiro.")

    def _traduzir(self, query: str) -> str:
        """Troca `?` pelo placeholder do driver, ignorando `?` dentro de aspas."""
        if self._get_placeholder() == "?":
            return query
        saida, aspa = [], None
        for ch in query:
            if aspa:
                if ch == aspa:
                    aspa = None
            elif ch in ("'", '"'):
                aspa = ch
            elif ch == "?":
                saida.append("%s")
                continue
            saida.append(ch)
        return "".join(saida)

    def _rollback(self) -> None:
        try:
            self._conn.rollback()
        except Exception:
            pass  # conexão caída: o erro que importa é o da query

    @staticmethod
    def _fechar(cursor) -> None:
        try:
            cursor.close()
        except Exception:
            pass

    def _execute(self, query: str, params: Optional[list] = None, as_dict: bool = False) -> list:
        self._garantir_conexao()
        cursor = self._conn.cursor()
        try:
            if params:
                cursor.execute(self._traduzir(query), list(params))
            else:
                # Sem parâmetros NÃO passa lista vazia: o pymysql aplicaria `%` no texto e um LIKE '%x%' estourava.
                cursor.execute(query)
            # description é None quando a instrução não devolve linhas (INSERT/UPDATE/DELETE/DDL).
            if cursor.description is not None:
                colunas = [c[0] for c in cursor.description]
                linhas = cursor.fetchall()
                results = [dict(zip(colunas, r)) for r in linhas] if as_dict else [tuple(r) for r in linhas]
            else:
                results = []
            self.last_rowcount = cursor.rowcount
            if not self._em_transacao:
                self._conn.commit()
            return results
        except Exception as e:
            self._rollback()
            raise QueryError(f"Erro ao executar query: {e}") from e
        finally:
            self._fechar(cursor)
