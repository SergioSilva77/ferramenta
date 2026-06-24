"""Módulo SQL para rpaflow."""

from typing import Any, Optional


class SQL:
    """Classe para operações com bancos de dados SQL."""

    def __init__(
        self,
        host: str = "localhost",
        user: str = "root",
        password: str = "",
        database: str = "",
        port: int = 3306,
        type: str = "mysql",
    ):
        self.host = host
        self.user = user
        self.password = password
        self.database = database
        self.port = port
        self.type = type.lower()
        self._conn = None

    def connect(self) -> bool:
        """Conecta ao banco de dados."""
        try:
            if self.type == "mysql":
                import pymysql
                self._conn = pymysql.connect(
                    host=self.host,
                    user=self.user,
                    password=self.password,
                    database=self.database,
                    port=self.port,
                )
            else:
                raise ValueError(f"Tipo '{self.type}' não suportado. Use 'mysql'.")
            return True
        except Exception as e:
            raise ConnectionError(f"Erro ao conectar: {e}")

    def disconnect(self) -> None:
        """Desconecta do banco de dados."""
        if self._conn:
            self._conn.close()
            self._conn = None

    def select(self, table: str, where: Optional[dict] = None) -> list:
        """Seleciona registros de uma tabela."""
        query = f"SELECT * FROM {table}"
        params = []

        if where:
            conditions = []
            for key, value in where.items():
                conditions.append(f"{key} = %s")
                params.append(value)
            query += " WHERE " + " AND ".join(conditions)

        return self._execute(query, params)

    def insert(self, table: str, data: dict) -> bool:
        """Insere um registro na tabela."""
        columns = ", ".join(data.keys())
        placeholders = ", ".join(["%s"] * len(data))
        query = f"INSERT INTO {table} ({columns}) VALUES ({placeholders})"
        self._execute(query, list(data.values()))
        return True

    def update(self, table: str, data: dict, where: dict) -> bool:
        """Atualiza registros na tabela."""
        set_clause = ", ".join([f"{k} = %s" for k in data.keys()])
        where_clause = " AND ".join([f"{k} = %s" for k in where.keys()])
        query = f"UPDATE {table} SET {set_clause} WHERE {where_clause}"
        params = list(data.values()) + list(where.values())
        self._execute(query, params)
        return True

    def delete(self, table: str, where: dict) -> bool:
        """Deleta registros da tabela."""
        where_clause = " AND ".join([f"{k} = %s" for k in where.keys()])
        query = f"DELETE FROM {table} WHERE {where_clause}"
        self._execute(query, list(where.values()))
        return True

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
