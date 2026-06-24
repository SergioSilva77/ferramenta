# JSON

Leitura e escrita de arquivos JSON com dot notation. Inspirado no Newtonsoft.Json (C#).

## Instalar

```bash
pip install rpaflow[json]
```

Ou sem instalar python-box (funciona com dict comum):

```python
from rpaflow.json import Json  # Funciona sem dependências extras
```

## Exemplo Rápido

```python
from rpaflow.json import Json

json = Json()
data = json.load("config.json")

# Dot notation
print(data.user.name)                # "Maria"
print(data.user.address.city)        # "SP"

# Iterar arrays
for order in data.orders:
    print(f"Order {order.id}: {order.total}")
```

## Exemplo Completo

```python
from rpaflow.json import Json

json = Json()

# ====== CARREGAR ======
data = json.load("C:/config/app.json")
data = json.loads('{"user": "Maria", "age": 30}')

# ====== DOT NOTATION ======
data.user.name                               # "Maria"
data.user.address.city                       # "SP"
data.bots[0].name                            # "invoice-bot"

# ====== ITERAR ARRAYS ======
# Iterar objetos em array
for bot in data.bots:
    print(f"Bot: {bot.name}, Enabled: {bot.enabled}")

# Iterar dicionário
for empresa, config in data.empresas.items():
    print(f"{empresa}: {config.host}:{config.port}")

# ====== ACESSAR/SETAR ======
data.user.name = "João"
data.user.phone = "+55 11 99999-0000"
data.user.address.country = "Brazil"         # Cria automaticamente

# ====== SALVAR ======
json.save("config.json", data)
json.save("config.json", data, indent=4)

# ====== SERIALIZAR ======
text = json.dumps(data)                      # Para string
text = json.dumps(data, indent=4)            # Formatado

# ====== UTILITÁRIOS ======
keys = json.get_keys(data.user)              # ["name", "address", "phone"]
values = json.get_values(data.user)          # ["Maria", {...}, "+55..."]
items = json.get_items(data.user)            # [("name", "Maria"), ...]
exists = json.has_key(data.user, "name")     # True
name = json.get_value(data.user, "name", "Default")  # "Maria"

# Converter de volta
plain_dict = json.to_dict(data)
plain_list = json.to_list(data.bots)
```

## Métodos

### Carregar

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `load()` | filepath, encoding | Carrega arquivo JSON |
| `loads()` | text | Parse string JSON |

### Salvar

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `save()` | filepath, data, indent, encoding | Salva em arquivo JSON |

### Serializar

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `dumps()` | data, indent | Serializa para string JSON |

### Utilitários

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get_keys()` | data | Lista de chaves |
| `get_values()` | data | Lista de valores |
| `get_items()` | data | Lista de tuplas (chave, valor) |
| `has_key()` | data, key | Verifica se chave existe |
| `get_value()` | data, key, default | Valor com fallback |
| `to_dict()` | data | Converte para dict |
| `to_list()` | data | Converte para list |

### Parâmetros Comuns

| Parâmetro | Padrão | Descrição |
|-----------|--------|-----------|
| `filepath` | — | Caminho do arquivo JSON |
| `encoding` | `"utf-8"` | Encoding do arquivo |
| `indent` | `2` | Indentação do JSON |
| `data` | — | Box, dict ou list |
| `key` | — | Nome da chave |
| `default` | `None` | Valor padrão |

## Dot Notation

### Acessar valores

```python
data = json.load("config.json")

# Com dot notation (igual Newtonsoft.Json)
data.user.name                       # "Maria"
data.user.address.city               # "SP"
data.database.host                   # "localhost"
data.bots[0].name                    # "invoice-bot"
data.bots[0].enabled                 # True

# Com colchetes (também funciona)
data["user"]["name"]                 # "Maria"
data["bots"][0]["name"]              # "invoice-bot"
```

### Setar valores

```python
data.user.name = "João"
data.user.phone = "+55 11 99999-0000"
data.user.address.country = "Brazil"  # Cria automaticamente
```

### Iterar arrays

```python
# Array de objetos
for bot in data.bots:
    print(f"Bot: {bot.name}")
    print(f"  Enabled: {bot.enabled}")
    print(f"  Retry: {bot.retry}")

# Dicionário
for empresa, config in data.empresas.items():
    print(f"{empresa}: {config.host}:{config.port}")
```

## Formato do Arquivo .json

```json
{
    "empresas": {
        "empresa1": {
            "host": "localhost",
            "port": 5432,
            "database": "vendas"
        },
        "empresa2": {
            "host": "192.168.1.100",
            "port": 3306,
            "database": "estoque"
        }
    },
    "bots": [
        {"name": "invoice-bot", "enabled": true, "retry": 3},
        {"name": "email-bot", "enabled": false, "retry": 5},
        {"name": "report-bot", "enabled": true, "retry": 1}
    ],
    "user": {
        "name": "Maria",
        "address": {
            "city": "São Paulo",
            "zip": "01000-000"
        }
    }
}
```

## Exemplos Práticos

### Iterar empresas e conectar ao banco

```python
from rpaflow.json import Json
from rpaflow.sql import SQL
from rpaflow.log import Log

json = Json()
log = Log(path="C:/logs/robô.log", level="DEBUG")

data = json.load("C:/config/empresas.json")

for empresa, config in data.empresas.items():
    log.info(f"Processando empresa: {empresa}")
    
    sql = SQL(
        type="mysql",
        host=config.host,
        user=config.user,
        password=config.password,
        database=config.database
    )
    sql.connect()
    
    pedidos = sql.select("pedidos", where={"status": "pendente"})
    log.info(f"Encontrados {len(pedidos)} pedidos")
    
    sql.disconnect()
```

### Iterar bots e executar

```python
from rpaflow.json import Json
from rpaflow.log import Log

json = Json()
log = Log(path="C:/logs/bots.log", level="DEBUG")

data = json.load("C:/config/bots.json")

for bot in data.bots:
    if bot.enabled:
        log.info(f"Executando bot: {bot.name}")
        # ... executar bot
        log.success(f"Bot {bot.name} finalizado")
    else:
        log.warning(f"Bot {bot.name} desabilitado")
```

### Carregar configuração de robô

```python
from rpaflow.json import Json
from rpaflow.sql import SQL
from rpaflow.log import Log
from rpaflow.excel_com import ExcelCom

json = Json()
data = json.load("C:/config/robô.json")

# Configurações do banco
db = data.database
sql = SQL(
    type=db.type,
    host=db.host,
    user=db.user,
    password=db.password,
    database=db.database
)

# Configurações de log
log_config = data.logging
log = Log(
    path=log_config.file,
    level=log_config.level,
    json=log_config.json
)

# Configurações do Excel
excel_config = data.excel
xl = ExcelCom(visible=excel_config.visible)
xl.open(excel_config.path)

log.info("Robô iniciado")
sql.connect()
# ... processar
log.success("Robô finalizado")
```

### Salvar dados processados

```python
from rpaflow.json import Json

json = Json()

# Criar dados
resultado = {
    "processados": 150,
    "erros": 3,
    "tempo_execucao": "00:05:30",
    "detalhes": [
        {"pedido": 1001, "status": "ok"},
        {"pedido": 1002, "status": "ok"},
        {"pedido": 1003, "status": "erro", "motivo": "timeout"}
    ]
}

# Salvar
json.save("C:/logs/resultado.json", resultado, indent=4)
print("Resultado salvo")
```

### Atualizar configuração existente

```python
from rpaflow.json import Json

json = Json()
data = json.load("C:/config/app.json")

# Atualizar valores
data.database.host = "192.168.1.100"
data.database.port = 5432

# Adicionar novo campo
data.database.timeout = 30

# Salvar
json.save("C:/config/app.json", data)
print("Configuração atualizada")
```

## Comparação com C# Newtonsoft.Json

| C# Newtonsoft.Json | Python rpaflow.json |
|--------------------|---------------------|
| `JsonConvert.DeserializeObject<T>(json)` | `json.loads(text)` |
| `JsonConvert.SerializeObject(obj)` | `json.dumps(data)` |
| `File.ReadAllText("config.json")` | `json.load("config.json")` |
| `File.WriteAllText("config.json", json)` | `json.save("config.json", data)` |
| `obj.User.Name` | `data.user.name` |
| `obj.Bots[0].Name` | `data.bots[0].name` |
| `foreach (var bot in obj.Bots)` | `for bot in data.bots:` |
| `foreach (var kvp in obj.Empresas)` | `for empresa, config in data.empresas.items():` |
