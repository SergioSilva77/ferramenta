# 05 - Sistema de Plugins

## Interface Base

Todo plugin herda `RPAPlugin`:

```python
from plugins.base import RPAPlugin

class MeuPlugin(RPAPlugin):
    @property
    def name(self):
        return "meu_plugin"

    @property
    def commands(self):
        return {
            "comando1": self.comando1,
            "comando2": self.comando2,
        }

    def comando1(self, **args):
        # implementação
        pass
```

## Como Adicionar um Novo Comando

1. Criar arquivo `plugins/novo_plugin.py`
2. Herdar de `RPAPlugin`
3. Declarar comandos no dict `commands`
4. Colocar arquivo na pasta `plugins/`
5. Pronto — PluginManager carrega automaticamente

**Não precisa tocar na grammar, no interpreter, nem gerar código ANTLR novamente.**

## Plugins Planejados

| Plugin | Comandos | Status |
|--------|----------|--------|
| `core_commands.py` | log, wait, sleep | Implementar primeiro |
| `file_plugin.py` | readFile, writeFile, copyFile, deleteFile | Futuro |
| `browser_plugin.py` | startBrowser, click, typeText, getText, screenshot, closeBrowser | Futuro |
| `excel_plugin.py` | openExcel, readSheet, writeCell, saveExcel, runMacro | Futuro |
| `database_plugin.py` | connectDB, query, disconnectDB | Futuro |
| `email_plugin.py` | sendEmail | Futuro |
| `api_plugin.py` | httpGet, httpPost, httpPut, httpDelete | Futuro |
