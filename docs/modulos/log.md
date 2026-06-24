# Log

Logging profissional inspirado no Serilog (C#). Usa loguru.

## Instalar

```bash
pip install rpaflow[log]
```

## Exemplo Rápido

```python
from rpaflow.log import Log

log = Log(path="C:/logs/meu_bot.log", level="DEBUG", json=True)
log.info("Iniciando robô")
log.success("Robô finalizado")
```

## Exemplo Completo

```python
from rpaflow.log import Log

# ====== CONFIGURAR ======
log = Log(
    path="C:/logs/meu_bot.log",   # Caminho do arquivo
    level="DEBUG",                  # Nível mínimo
    rotation="10 MB",               # Rotação por tamanho
    retention="7 days",             # Manter 7 dias
    compression="zip",              # Comprimir logs antigos
    json=True,                      # Output JSON no arquivo
    console=True,                   # Exibir no console
)

# ====== NÍVEIS DE LOG ======
log.debug("Depuração")
log.info("Informação")
log.warning("Aviso")
log.error("Erro")
log.critical("Crítico")
log.success("Sucesso")
log.exception("Exception com traceback")

# ====== COM CONTEXTO (kwargs) ======
log.info("Conectando ao banco", host="localhost", user="admin")
log.error("Falha na conexão", host="localhost", timeout=30)

# ====== COM BIND (tipo ForContext do Serilog) ======
log = log.bind(user="admin", bot="vendas")
log.info("Processando pedido", pedido_id=123)
# Output: 2026-06-24 10:30:15 | INFO     | user=admin | bot=vendas | Processando pedido | pedido_id=123

# ====== EXCEPTION ======
try:
    1 / 0
except:
    log.exception("Erro inesperado")
```

## Métodos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `debug()` | message, **kwargs | Log DEBUG |
| `info()` | message, **kwargs | Log INFO |
| `warning()` | message, **kwargs | Log WARNING |
| `error()` | message, **kwargs | Log ERROR |
| `critical()` | message, **kwargs | Log CRITICAL |
| `success()` | message, **kwargs | Log SUCCESS |
| `exception()` | message, **kwargs | Log EXCEPTION com traceback |
| `bind()` | **kwargs | Adiciona contexto (tipo ForContext) |

## Parâmetros do Construtor

| Parâmetro | Padrão | Descrição |
|-----------|--------|-----------|
| `path` | `None` | Caminho do arquivo (None = só console) |
| `level` | `"DEBUG"` | Nível mínimo (DEBUG, INFO, WARNING, ERROR, CRITICAL) |
| `rotation` | `"10 MB"` | Rotação ("10 MB", "00:00", "1 week") |
| `retention` | `"7 days"` | Manter logs antigos ("7 days", "30 days") |
| `compression` | `"zip"` | Compressão ("zip", "gz", "tar") |
| `json` | `False` | Output JSON no arquivo |
| `console` | `True` | Exibir no console |
| `format_console` | Auto | Formato customizado do console |
| `format_file` | Auto | Formato customizado do arquivo |

## Formatos de Rotação

| Valor | Descrição |
|-------|-----------|
| `"10 MB"` | Rotação por tamanho (10 megabytes) |
| `"100 KB"` | Rotação por tamanho (100 kilobytes) |
| `"1 GB"` | Rotação por tamanho (1 gigabyte) |
| `"00:00"` | Rotação à meia-noite |
| `"1 week"` | Rotação semanal |
| `"30 days"` | Rotação mensal |

## Formatos de Retenção

| Valor | Descrição |
|-------|-----------|
| `"7 days"` | Manter 7 dias |
| `"30 days"` | Manter 30 dias |
| `"1 year"` | Manter 1 ano |
| `"10"` | Manter 10 arquivos |

## Formatos de Compressão

| Valor | Descrição |
|-------|-----------|
| `"zip"` | Compressão ZIP |
| `"gz"` | Compressão GZIP |
| `"tar"` | Compressão TAR |
| `"tar.gz"` | Compressão TAR.GZ |

## Retorno no Console

```
2026-06-24 10:30:15 | DEBUG    | Depuração
2026-06-24 10:30:15 | INFO     | Informação
2026-06-24 10:30:16 | WARNING  | Aviso
2026-06-24 10:30:16 | ERROR    | Erro
2026-06-24 10:30:16 | CRITICAL | Crítico
2026-06-24 10:30:16 | SUCCESS  | Sucesso
```

## Retorno no Arquivo JSON

```json
{"text": "2026-06-24 10:30:15", "level": "INFO", "message": "Informação"}
{"text": "2026-06-24 10:30:16", "level": "ERROR", "message": "Erro"}
{"text": "2026-06-24 10:30:16", "level": "ERROR", "message": "Falha na conexão | host=localhost | timeout=30"}
```

## Retorno com Bind

```python
log = log.bind(user="admin", bot="vendas")
log.info("Processando pedido", pedido_id=123)
```

```
2026-06-24 10:30:15 | INFO     | user=admin | bot=vendas | Processando pedido | pedido_id=123
```

## Exemplo com Robô

```python
from rpaflow.log import Log
from rpaflow.sql import SQL
from rpaflow.excel_com import ExcelCom

# Configurar log
log = Log(
    path="C:/logs/vendas.log",
    level="DEBUG",
    rotation="10 MB",
    retention="7 days",
    json=True,
)

# Contexto do robô
log = log.bind(bot="vendas", user="admin")

log.info("Iniciando robô de vendas")

# Conectar ao banco
sql = SQL(type="mysql", host="localhost", user="root", password="123", database="vendas")
sql.connect()
log.info("Conectado ao banco")

# Buscar pedidos
pedidos = sql.select("pedidos", where={"status": "pendente"})
log.info("Pedidos encontrados", total=len(pedidos))

# Processar
for pedido in pedidos:
    try:
        # ... processar pedido
        log.success("Pedido processado", pedido_id=pedido[0])
    except Exception as e:
        log.error("Falha no pedido", pedido_id=pedido[0], error=str(e))

log.success("Robô finalizado com sucesso")
```
