"""Módulo SQL para rpaflow."""

from typing import Any, Optional


class SQL:
    """Classe para operações com bancos de dados SQL.

    Tipos suportados: mysql, postgresql, sqlserver, sqlite
    """

    def __init__(
        self,
        host: str = "localhost",
        user: str = "",
        password: str = "",
        database: str = "",
        port: int = 0,
        type: str = "mysql",
    ):
        self.host = host
        self.user = user
        self.password = password
        self.database = database
        self.type = type.lower()

        # Portas padrão por tipo
        default_ports = {
            "mysql": 3306,
            "postgresql": 5432,
            "sqlserver": 1433,
            "sqlite": 0,
        }
        self.port = port or default_ports.get(self.type, 3306)
        self._conn = None

    def connect(self) -> bool:
        """Conecta ao banco de dados."""
        try:
            self._conn = self._get_connection()
            return True
        except Exception as e:
            raise ConnectionError(f"Erro ao conectar: {e}")

    def _get_connection(self):
        """Cria a conexão baseado no tipo."""
        if self.type == "mysql":
            import pymysql
            return pymysql.connect(
                host=self.host,
                user=self.user,
                password=self.password,
                database=self.database,
                port=self.port,
            )

        elif self.type == "postgresql":
            import psycopg2
            return psycopg2.connect(
                host=self.host,
                user=self.user,
                password=self.password,
                dbname=self.database,
                port=self.port,
            )

        elif self.type == "sqlserver":
            import pyodbc
            conn_str = (
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={self.host},{self.port};"
                f"DATABASE={self.database};"
                f"UID={self.user};"
                f"PWD={self.password}"
            )
            return pyodbc.connect(conn_str)

        elif self.type == "sqlite":
            import sqlite3
            return sqlite3.connect(self.database)

        else:
            raise ValueError(
                f"Tipo '{self.type}' não suportado. "
                "Use: mysql, postgresql, sqlserver ou sqlite."
            )

    def _get_placeholder(self) -> str:
        """Retorna o placeholder correto para o tipo de banco."""
        if self.type in ("mysql", "postgresql"):
            return "%s"
        else:  # sqlserver, sqlite
            return "?"

    def disconnect(self) -> None:
        """Desconecta do banco de dados."""
        if self._conn:
            self._conn.close()
            self._conn = None

    def select(self, table: str, where: Optional[dict] = None) -> list:
        """Seleciona registros de uma tabela."""
        ph = self._get_placeholder()
        query = f"SELECT * FROM {table}"
        params = []

        if where:
            conditions = []
            for key, value in where.items():
                conditions.append(f"{key} = {ph}")
                params.append(value)
            query += " WHERE " + " AND ".join(conditions)

        return self._execute(query, params)

    def insert(self, table: str, data: dict) -> bool:
        """Insere um registro na tabela."""
        ph = self._get_placeholder()
        columns = ", ".join(data.keys())
        placeholders = ", ".join([ph] * len(data))
        query = f"INSERT INTO {table} ({columns}) VALUES ({placeholders})"
        self._execute(query, list(data.values()))
        return True

    def update(self, table: str, data: dict, where: dict) -> bool:
        """Atualiza registros na tabela."""
        ph = self._get_placeholder()
        set_clause = ", ".join([f"{k} = {ph}" for k in data.keys()])
        where_clause = " AND ".join([f"{k} = {ph}" for k in where.keys()])
        query = f"UPDATE {table} SET {set_clause} WHERE {where_clause}"
        params = list(data.values()) + list(where.values())
        self._execute(query, params)
        return True

    def delete(self, table: str, where: dict) -> bool:
        """Deleta registros da tabela."""
        ph = self._get_placeholder()
        where_clause = " AND ".join([f"{k} = {ph}" for k in where.keys()])
        query = f"DELETE FROM {table} WHERE {where_clause}"
        self._execute(query, list(where.values()))
        return True

    def execute(self, query: str, params: list = None) -> list:
        """Executa uma query customizada e retorna os resultados."""
        return self._execute(query, params)

    def _execute(self, query: str, params: list = None) -> list:
        """Executa uma query e retorna os resultados."""
        if not self._conn:
            raise ConnectionError("Não conectado ao banco. Chame connect() primeiro.")

        try:
            cursor = self._conn.cursor()
            cursor.execute(query, params or [])
            results = cursor.fetchall()
            self._conn.commit()
            return results
        except Exception as e:
            self._conn.rollback()
            raise QueryError(f"Erro ao executar query: {e}")
