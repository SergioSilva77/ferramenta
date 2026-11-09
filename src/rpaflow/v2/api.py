"""Módulo API v2 para rpaflow — requisições HTTP/REST.

Diferenças em relação ao `rpaflow.api.API` (v1), as principais:

- `APIError` agora é importado (no v1 qualquer falha virava `NameError`).
- `params`, `data`, `timeout` e `follow_redirects` (o v1 não seguia redirect e só aceitava JSON no corpo).
- `raise_for_status=True` opcional: HTTP 4xx/5xx vira `APIError` em vez de passar batido como "sucesso".
- `patch`, `request` genérico, `base_url`, headers padrão, `with API() as api:`.
"""

from __future__ import annotations

from typing import Optional

from rpaflow._exceptions import APIError


class API:
    """Classe para requisições HTTP/REST (httpx)."""

    def __init__(
        self,
        base_url: str = "",
        headers: Optional[dict] = None,
        timeout: float = 30.0,
        raise_for_status: bool = False,
        follow_redirects: bool = True,
        verify: bool = True,
    ):
        """
        Args:
            base_url: prefixo para URLs relativas (`api.get("/clientes")`).
            headers: headers enviados em toda requisição (ex: Authorization).
            timeout: segundos. O padrão do httpx é 5s, curto demais para portais lentos.
            raise_for_status: levanta `APIError` em resposta 4xx/5xx.
            verify: False ignora erro de certificado (só para ambiente interno com cert próprio).
        """
        self.base_url = base_url
        self.headers = dict(headers or {})
        self.timeout = timeout
        self.raise_for_status = raise_for_status
        self.follow_redirects = follow_redirects
        self.verify = verify
        self._client = None

    def _get_client(self):
        if not self._client:
            import httpx
            self._client = httpx.Client(
                base_url=self.base_url,
                headers=self.headers,
                timeout=self.timeout,
                follow_redirects=self.follow_redirects,
                verify=self.verify,
            )
        return self._client

    def request(
        self,
        method: str,
        url: str,
        *,
        params: Optional[dict] = None,
        json: Optional[object] = None,
        data: Optional[object] = None,
        headers: Optional[dict] = None,
        files: Optional[dict] = None,
        timeout: Optional[float] = None,
    ):
        """Requisição genérica. Devolve o `httpx.Response`."""
        try:
            extra = {"timeout": timeout} if timeout is not None else {}
            resp = self._get_client().request(
                method.upper(), url, params=params, json=json, data=data,
                headers=headers, files=files, **extra,
            )
        except Exception as e:
            raise APIError(f"Erro na requisição {method.upper()} {url}: {e}") from e
        if self.raise_for_status and resp.status_code >= 400:
            raise APIError(f"HTTP {resp.status_code} em {method.upper()} {url}: {resp.text[:300]}")
        return resp

    def get(self, url: str, headers: Optional[dict] = None, params: Optional[dict] = None, **kw):
        """GET."""
        return self.request("GET", url, headers=headers, params=params, **kw)

    def post(self, url: str, json: Optional[object] = None, headers: Optional[dict] = None,
             data: Optional[object] = None, **kw):
        """POST (json= para corpo JSON, data= para formulário)."""
        return self.request("POST", url, json=json, headers=headers, data=data, **kw)

    def put(self, url: str, json: Optional[object] = None, headers: Optional[dict] = None,
            data: Optional[object] = None, **kw):
        """PUT."""
        return self.request("PUT", url, json=json, headers=headers, data=data, **kw)

    def patch(self, url: str, json: Optional[object] = None, headers: Optional[dict] = None,
              data: Optional[object] = None, **kw):
        """PATCH."""
        return self.request("PATCH", url, json=json, headers=headers, data=data, **kw)

    def delete(self, url: str, headers: Optional[dict] = None, params: Optional[dict] = None, **kw):
        """DELETE."""
        return self.request("DELETE", url, headers=headers, params=params, **kw)

    def close(self) -> None:
        """Fecha o cliente HTTP."""
        if self._client:
            self._client.close()
            self._client = None

    def __enter__(self) -> "API":
        return self

    def __exit__(self, *exc) -> bool:
        self.close()
        return False
