"""Módulo API para rpaflow."""

from typing import Optional


class API:
    """Classe para requisições HTTP/REST."""

    def __init__(self):
        self._client = None

    def _get_client(self):
        """Retorna ou cria o cliente HTTP."""
        if not self._client:
            import httpx
            self._client = httpx.Client()
        return self._client

    def get(self, url: str, headers: Optional[dict] = None):
        """Realiza uma requisição GET."""
        try:
            client = self._get_client()
            return client.get(url, headers=headers or {})
        except Exception as e:
            raise APIError(f"Erro na requisição GET: {e}")

    def post(self, url: str, json: Optional[dict] = None, headers: Optional[dict] = None):
        """Realiza uma requisição POST."""
        try:
            client = self._get_client()
            return client.post(url, json=json, headers=headers or {})
        except Exception as e:
            raise APIError(f"Erro na requisição POST: {e}")

    def put(self, url: str, json: Optional[dict] = None, headers: Optional[dict] = None):
        """Realiza uma requisição PUT."""
        try:
            client = self._get_client()
            return client.put(url, json=json, headers=headers or {})
        except Exception as e:
            raise APIError(f"Erro na requisição PUT: {e}")

    def delete(self, url: str, headers: Optional[dict] = None):
        """Realiza uma requisição DELETE."""
        try:
            client = self._get_client()
            return client.delete(url, headers=headers or {})
        except Exception as e:
            raise APIError(f"Erro na requisição DELETE: {e}")

    def close(self) -> None:
        """Fecha o cliente HTTP."""
        if self._client:
            self._client.close()
            self._client = None
