# API

Requisições HTTP/REST.

## Instalar

```bash
pip install rpaflow[api]
```

## Exemplo

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

## Métodos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get()` | url, headers (opcional) | Requisição GET |
| `post()` | url, json, headers (opcional) | Requisição POST |
| `put()` | url, json, headers (opcional) | Requisição PUT |
| `delete()` | url, headers (opcional) | Requisição DELETE |
