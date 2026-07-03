# Browser

Automação de navegador web (Playwright ou Selenium).

## Instalar

```bash
pip install rpaflow[browser]
```

## Exemplo Rápido

```python
from rpaflow.browser import Browser

browser = Browser()
browser.start("https://site.com", type="playwright")
browser.click("#botao")
browser.close()
```

## Exemplo Completo

```python
from rpaflow.browser import Browser

browser = Browser()
browser.start("https://site.com", type="playwright")

# ====== NAVEGAÇÃO ======
browser.navigate("https://site.com/login")
browser.wait_for_load_state("networkidle")

# ====== INTERAÇÃO ======
browser.type_text("#user", "admin", delay=100)
browser.fill_text("#pass", "123456")
browser.click("#btn-login")
browser.select_option("#dropdown", "Opção 1")
browser.hover("#menu-item")

# ====== TEXTO ======
texto = browser.get_text("h1")
items = browser.get_all_elements(".item")

# ====== ESPERAR ======
browser.wait_for_element("#dashboard", state="visible")

# ====== TECLADO ======
browser.send_shortcut("Enter")
browser.send_shortcut("Control+A")
browser.send_shortcut("Escape")

# ====== FRAME (explícito) ======
browser.click("#botao", frame="#iframe-1")
browser.select_option("#dropdown", "Opção 1", frame="#iframe-1")
browser.type_text("#campo", "texto", frame="#iframe-1")

# ====== RECURSIVE (busca em todos os iframes) ======
browser.click("#botao", recursive=True)
browser.select_option("#dropdown", "Opção 1", recursive=True)
browser.type_text("#campo", "texto", recursive=True)
browser.get_text("#texto", recursive=True)

# ====== DOWNLOAD ======
filepath = browser.wait_for_download("C:/downloads/arquivo.pdf")

# ====== JANELA ======
browser.maximize()

# ====== ABAS ======
browser.new_tab("https://google.com")
browser.new_tab("https://github.com")

tabs = browser.get_tabs()
print(tabs)
# [{'index': 0, 'title': 'Site', 'url': 'https://site.com'},
#  {'index': 1, 'title': 'Google', 'url': 'https://google.com'},
#  {'index': 2, 'title': 'GitHub', 'url': 'https://github.com'}]

browser.switch_to_tab(0)      # volta para a primeira aba
browser.close_tab(2)          # fecha a aba do GitHub

# ====== SCREENSHOT ======
browser.screenshot("captura.png")

# ====== FECHAR ======
browser.close()
```

## Métodos

### Iniciar / Fechar

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `start()` | url, type, headless | Inicia navegador |
| `close()` | — | Fecha navegador |

### Navegação

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `navigate()` | url | Navega para URL |
| `wait_for_load_state()` | state, timeout | Espera carregamento |

### Clique

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `click()` | selector, frame, recursive, timeout | Clica em elemento |
| `hover()` | selector, frame, recursive, timeout | Passa mouse sobre elemento |

### Texto

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `type_text()` | selector, text, delay, frame, recursive, timeout | Digita caractere por caractere |
| `fill_text()` | selector, text, frame, recursive, timeout | Preenche instantaneamente |
| `get_text()` | selector, frame, recursive, timeout | Retorna texto visível |
| `get_all_elements()` | selector, frame, recursive, timeout | Retorna lista de elementos |

### Select

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `select_option()` | selector, label, frame, recursive, timeout | Seleciona por texto |
| `select_option_by_index()` | selector, index, frame, recursive, timeout | Seleciona por índice (1-based) |

### Esperar

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `wait_for_element()` | selector, state, frame, recursive, timeout | Espera estado do elemento |
| `wait_for_download()` | save_path, timeout | Espera download concluir |

### Teclado

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `send_shortcut()` | key | Envia atalho de teclado |

### Frame

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `switch_to_frame()` | selector | Entra no iframe |
| `switch_to_root()` | — | Sai do iframe |
| `find_element_in_frames()` | selector, timeout | Busca elemento em todos os iframes |
| `click_in_frames()` | selector, timeout | Clica em qualquer iframe |
| `select_in_frames()` | selector, label, timeout | Seleciona opção em qualquer iframe |
| `find_all_in_frames()` | selector, timeout | Busca múltiplos elementos em qualquer iframe |

### Property

| Propriedade | Descrição |
|-------------|-----------|
| `page` | Retorna o Playwright Page subjacente |

### Janela

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `maximize()` | — | Maximiza a janela usando dimensões reais da tela |

### Abas

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get_tabs()` | — | Lista abas com index, title e url |
| `get_current_tab()` | — | Retorna info da aba atual |
| `switch_to_tab()` | index (0-based) | Muda para a aba pelo índice |
| `new_tab()` | url (opcional) | Abre nova aba |
| `close_tab()` | index (opcional, 0-based) | Fecha aba (None = fecha atual) |

### Screenshot

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `screenshot()` | filepath | Captura screenshot |

## Parâmetros Comuns

| Parâmetro | Padrão | Descrição |
|-----------|--------|-----------|
| `selector` | — | Seletor CSS, XPath, etc. |
| `frame` | `None` | Seletor do iframe (explícito) |
| `recursive` | `False` | Busca em todos os iframes |
| `timeout` | `5000` | Timeout em ms |
| `delay` | `50` | Delay entre teclas (ms) |
| `state` | `"visible"` | Estado: visible, hidden, attached, detached |

## Iframe: Frame vs Recursive

### Frame explícito (quando sabe qual iframe)

```python
browser.click("#botao", frame="#iframe-1")
```

### Recursive (busca em todos os iframes)

```python
browser.click("#botao", recursive=True)
browser.select_option("#dropdown", "Opção 1", recursive=True)
browser.type_text("#campo", "texto", recursive=True)
browser.get_text("#texto", recursive=True)
```

### Todos os métodos suportam `recursive=True`

| Método | Recursive |
|--------|:---------:|
| `click()` | Sim |
| `hover()` | Sim |
| `type_text()` | Sim |
| `fill_text()` | Sim |
| `get_text()` | Sim |
| `get_all_elements()` | Sim |
| `select_option()` | Sim |
| `select_option_by_index()` | Sim |
| `wait_for_element()` | Sim |

## Atalhos de Teclado

```python
browser.send_shortcut("Enter")
browser.send_shortcut("Control+A")
browser.send_shortcut("Control+C")
browser.send_shortcut("Control+V")
browser.send_shortcut("Escape")
browser.send_shortcut("Tab")
browser.send_shortcut("Shift+Tab")
```
