# 06 - Criando Plugins (Módulos Novos)

## Como Funciona

Cada módulo do rpaflow é uma pasta com `__init__.py`. Para criar um novo módulo, basta criar uma nova pasta dentro de `rpaflow/`.

## Estrutura de um Plugin

```
rpaflow/
└── src/
    └── rpaflow/
        ├── sql/                    # Módulo built-in
        │   ├── __init__.py
        │   └── _client.py
        │
        ├── meu_plugin/             # Seu novo módulo
        │   ├── __init__.py         # Expõe os métodos públicos
        │   └── _interno.py         # Lógica interna (opcional)
        │
        └── outro_plugin/
            ├── __init__.py
            └── _core.py
```

## Criando um Plugin do Zero

### Passo 1: Criar a pasta

```
src/rpaflow/sap/
├── __init__.py
└── _client.py
```

### Passo 2: Implementar o módulo interno

```python
# src/rpaflow/sap/_client.py

class SAPClient:
    def __init__(self):
        self.conn = None

    def connect(self, host, user, password, system):
        # Lógica de conexão SAP
        print(f"Conectando ao SAP {system} em {host}")
        self.conn = {"host": host, "user": user}
        return True

    def call_function(self, func_name, params=None):
        if not self.conn:
            raise Exception("SAP não conectado")
        # Lógica para chamar RFC
        return {"result": "ok", "function": func_name}

    def disconnect(self):
        self.conn = None
        print("Desconectado do SAP")
```

### Passo 3: Criar o `__init__.py` público

```python
# src/rpaflow/sap/__init__.py
"""Módulo SAP para rpaflow."""

from rpaflow.sap._client import SAPClient

# Instância única do cliente
_client = SAPClient()

# Métodos públicos (api limpa)
connect = _client.connect
call_function = _client.call_function
disconnect = _client.disconnect
```

### Passo 4: Usar

```python
from rpaflow import sap

sap.connect(host="10.0.0.1", user="admin", password="123", system="PRD")
result = sap.call_function("RFC_READ_TABLE", {"QUERY_TABLE": "T001"})
sap.disconnect()
```

---

## Exemplo: Plugin Salesforce

```
src/rpaflow/salesforce/
├── __init__.py
└── _client.py
```

```python
# src/rpaflow/salesforce/_client.py
import httpx

class SalesforceClient:
    def __init__(self):
        self.token = None
        self.instance_url = None

    def connect(self, username, password, security_token, domain="login"):
        url = f"https://{domain}.salesforce.com/services/oauth2/token"
        data = {
            "grant_type": "password",
            "client_id": "dummy",
            "client_secret": "dummy",
            "username": username,
            "password": password + security_token,
        }
        # Em produção: response = httpx.post(url, data=data)
        # self.token = response.json()["access_token"]
        # self.instance_url = response.json()["instance_url"]
        print(f"Conectado ao Salesforce: {domain}")
        return True

    def query(self, soql):
        if not self.token:
            raise Exception("Salesforce não conectado")
        # Em produção: response = httpx.get(...)
        return []

    def create(self, object_name, data):
        print(f"Criando {object_name}: {data}")
        return {"id": "001xx000003DGP0", "success": True}

    def update(self, object_name, record_id, data):
        print(f"Atualizando {object_name} {record_id}: {data}")
        return True

    def delete(self, object_name, record_id):
        print(f"Deletando {object_name} {record_id}")
        return True

    def disconnect(self):
        self.token = None
        self.instance_url = None
```

```python
# src/rpaflow/salesforce/__init__.py
"""Módulo Salesforce para rpaflow."""

from rpaflow.salesforce._client import SalesforceClient

_client = SalesforceClient()

connect = _client.connect
query = _client.query
create = _client.create
update = _client.update
delete = _client.delete
disconnect = _client.disconnect
```

**Uso:**

```python
from rpaflow import salesforce

salesforce.connect(
    username="user@email.com",
    password="senha",
    security_token="abc123"
)

contacts = salesforce.query("SELECT Id, Name FROM Contact LIMIT 10")

salesforce.create("Contact", {
    "FirstName": "João",
    "LastName": "Silva",
    "Email": "joao@email.com"
})

salesforce.disconnect()
```

---

## Exemplo: Plugin Interno (ERP)

```
src/rpaflow/erp/
├── __init__.py
└── _api.py
```

```python
# src/rpaflow/erp/_api.py
import httpx

class ERPClient:
    def __init__(self):
        self.base_url = None
        self.token = None

    def connect(self, url, token):
        self.base_url = url
        self.token = token
        print(f"Conectado ao ERP: {url}")
        return True

    def get_orders(self, status=None):
        # response = httpx.get(f"{self.base_url}/orders", headers={"Authorization": f"Bearer {self.token}"})
        return []

    def create_order(self, data):
        print(f"Criando pedido: {data}")
        return {"id": 1, "status": "created"}

    def update_order(self, order_id, data):
        print(f"Atualizando pedido {order_id}: {data}")
        return True

    def disconnect(self):
        self.base_url = None
        self.token = None
```

```python
# src/rpaflow/erp/__init__.py
"""Módulo ERP para rpaflow."""

from rpaflow.erp._api import ERPClient

_client = ERPClient()

connect = _client.connect
get_orders = _client.get_orders
create_order = _client.create_order
update_order = _client.update_order
disconnect = _client.disconnect
```

**Uso:**

```python
from rpaflow import erp

erp.connect(url="https://erp.empresa.com/api", token="abc123")

orders = erp.get_orders(status="pending")
print(f"{len(orders)} pedidos pendentes")

erp.create_order({"client": "João", "product": "Notebook", "qty": 1})

erp.disconnect()
```

---

## Resumo

| Ação | Como fazer |
|------|-----------|
| Criar módulo | Criar pasta `rpaflow/meu_modulo/` |
| Api pública | Definir métodos no `__init__.py` |
| Lógica interna | Arquivo `_interno.py` (com underscore) |
| Usar | `from rpaflow import meu_modulo` |
| Dependências | Adicionar no `pyproject.toml` sob `[project.optional-dependencies]` |

**Não precisa registrar nada.** Só criar a pasta e importar.
