# Browser

Automação de navegador web (Playwright ou Selenium).

## Instalar

```bash
pip install rpaflow[browser]
```

## Exemplo

```python
from rpaflow.browser import Browser

# Iniciar navegador
browser = Browser()
browser.start(url="https://site.com", type="playwright")

# Clicar
browser.click(selector="#botao")

# Digitar
browser.type(selector="#input", text="Olá")

# Ler texto
texto = browser.get_text(selector="h1")

# Screenshot
browser.screenshot("captura.png")

# Fechar
browser.close()
```

## Métodos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `start()` | url, type (playwright/selenium) | Abre navegador |
| `click()` | selector | Clica em elemento |
| `type()` | selector, text | Digita em campo |
| `get_text()` | selector | Retorna texto |
| `screenshot()` | filepath | Captura tela |
| `close()` | — | Fecha navegador |
