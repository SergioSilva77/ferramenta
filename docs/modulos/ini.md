# INI

Leitura e escrita de arquivos .ini (configurações).

## Instalar

```bash
pip install rpaflow[ini]
```

Ou sem instalar nada extra (configparser é built-in):

```python
from rpaflow.ini import Ini
```

## Exemplo Rápido

```python
from rpaflow.ini import Ini

ini = Ini("C:/config/app.ini")
host = ini.get("database", "host")
port = ini.get_int("database", "port")
```

## Exemplo Completo

```python
from rpaflow.ini import Ini

# ====== ABRIR ARQUIVO ======
ini = Ini("C:/config/app.ini")

# ====== SEÇÕES ======
sections = ini.get_sections()              # ["database", "logging", "app"]
ini.has_section("database")                # True

# ====== ITENS ======
items = ini.get_items("database")          # [("host", "localhost"), ("port", "5432")]
ini.has_item("database", "host")           # True

# ====== VALORES ======
host = ini.get("database", "host")                        # "localhost"
port = ini.get_int("database", "port")                     # 5432
debug = ini.get_bool("app", "debug")                       # True
timeout = ini.get_float("app", "timeout")                  # 30.0

# Com fallback
host = ini.get("database", "host", fallback="localhost")
port = ini.get_int("database", "port", fallback=3306)

# ====== ESCREVER ======
ini.set("database", "host", "192.168.1.100")
ini.set("database", "port", 3306)
ini.set("app", "debug", True)
ini.save()

# ====== ADICIONAR SEÇÃO ======
ini.add_section("cache")
ini.set("cache", "backend", "redis")
ini.set("cache", "ttl", 3600)
ini.save()

# ====== REMOVER ======
ini.remove_item("database", "password")
ini.remove_section("cache")
ini.save()

# ====== CRIAR DO ZERO ======
ini = Ini("C:/config/novo.ini")
ini.add_section("database")
ini.set("database", "host", "localhost")
ini.set("database", "port", 5432)
ini.add_section("logging")
ini.set("logging", "level", "INFO")
ini.save()

# ====== UTILITÁRIOS ======
ini.print_all()                            # Imprime tudo
dados = ini.to_dict()                      # Converte para dict
```

## Métodos

### Seções

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get_sections()` | — | Lista de seções |
| `has_section()` | section | Verifica se seção existe |
| `add_section()` | section | Adiciona seção |
| `remove_section()` | section | Remove seção |

### Itens

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get_items()` | section | Lista de tuplas (chave, valor) |
| `has_item()` | section, key | Verifica se item existe |

### Valores

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get()` | section, key, fallback | Retorna como string |
| `get_int()` | section, key, fallback | Retorna como int |
| `get_bool()` | section, key, fallback | Retorna como bool |
| `get_float()` | section, key, fallback | Retorna como float |

### Escrever

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `set()` | section, key, value | Define valor |
| `remove_item()` | section, key | Remove item |
| `save()` | filepath | Salva arquivo |

### Utilitários

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `print_all()` | — | Imprime seções e itens |
| `to_dict()` | — | Converte para dict |

### Parâmetros Comuns

| Parâmetro | Padrão | Descrição |
|-----------|--------|-----------|
| `section` | — | Nome da seção |
| `key` | — | Nome da chave |
| `fallback` | `None` | Valor padrão se não encontrar |
| `encoding` | `"utf-8"` | Encoding do arquivo |

## Formato do Arquivo .ini

```ini
[database]
host = localhost
port = 5432
user = admin
password = secret

[logging]
level = INFO
file = /var/log/app.log

[app]
debug = true
timeout = 30.0
name = MeuApp
```

## Exemplos Práticos

### Ler configuração de robô

```python
from rpaflow.ini import Ini
from rpaflow.sql import SQL
from rpaflow.log import Log

ini = Ini("C:/config/robô.ini")

db_host = ini.get("database", "host")
db_user = ini.get("database", "user")
db_pass = ini.get("database", "password")
db_name = ini.get("database", "database")
log_level = ini.get("logging", "level", fallback="INFO")

log = Log(path="C:/logs/robô.log", level=log_level)
sql = SQL(type="mysql", host=db_host, user=db_user, password=db_pass, database=db_name)

log.info("Robô iniciado", host=db_host)
sql.connect()
```

### Criar configuração padrão

```python
from rpaflow.ini import Ini

ini = Ini("C:/config/robô.ini")

if not ini.has_section("database"):
    ini.add_section("database")
    ini.set("database", "host", "localhost")
    ini.set("database", "port", 3306)
    ini.set("database", "user", "admin")
    ini.save()
    print("Configuração padrão criada")
```

### Atualizar valor existente

```python
from rpaflow.ini import Ini

ini = Ini("C:/config/app.ini")

# Atualizar porta
ini.set("database", "port", 5432)
ini.save()

# Verificar
port = ini.get_int("database", "port")
print(f"Porta atualizada: {port}")
```

### Listar todas as configurações

```python
from rpaflow.ini import Ini

ini = Ini("C:/config/app.ini")
ini.print_all()
```

```
[database]
host = localhost
port = 5432
user = admin

[logging]
level = INFO
file = /var/log/app.log

[app]
debug = true
timeout = 30.0
```

### Converter para dict

```python
from rpaflow.ini import Ini

ini = Ini("C:/config/app.ini")
config = ini.to_dict()

print(config)
# {
#     "database": {"host": "localhost", "port": "5432", "user": "admin"},
#     "logging": {"level": "INFO", "file": "/var/log/app.log"},
#     "app": {"debug": "true", "timeout": "30.0"},
# }
```
