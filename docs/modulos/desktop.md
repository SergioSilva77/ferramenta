# Desktop

Automação desktop via reconhecimento de imagem (pyautogui + opencv).

## Instalar

```bash
pip install rpaflow[desktop]
```

## Exemplo Rápido

```python
from rpaflow.desktop import Desktop

desktop = Desktop()
desktop.click_image("C:/imgs/botao.png")
```

## Exemplo Completo

```python
from rpaflow.desktop import Desktop

desktop = Desktop()

# ====== LOCALIZAR (sem clicar) ======
result = desktop.find_image("C:/imgs/botao.png")
if result:
    print(f"Encontrado em x={result['x']}, y={result['y']}")

# ====== CLICAR ======
# Clique simples
desktop.click_image("C:/imgs/botao.png")

# Com confiança
desktop.click_image("C:/imgs/botao.png", confidence=0.90)

# Com offset
desktop.click_image("C:/imgs/botao.png", offset="right", offset_x=10)

# Clicar na segunda ocorrência
desktop.click_image("C:/imgs/icone.png", match_index=1)

# Com highlight para debug
desktop.click_image("C:/imgs/botao.png", highlight_ms=2000)

# Ajustar tentativas
desktop.click_image("C:/imgs/botao.png", limit=20, time_ms=500)

# ====== ENCONTRAR TODAS ======
results = desktop.find_all_images("C:/imgs/icone.png")
for r in results:
    print(f"Ocorrência {r['index']}: x={r['x']}, y={r['y']}")
```

## Métodos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `find_image()` | image, confidence, limit, time_ms | Localiza imagem e retorna coordenadas |
| `click_image()` | image, confidence, offset, offset_x, offset_y, match_index, highlight_ms, limit, time_ms | Localiza e clica na imagem |
| `find_all_images()` | image, confidence, limit, time_ms | Localiza todas as ocorrências |

## Parâmetros

| Parâmetro | Padrão | Descrição |
|-----------|--------|-----------|
| `image` | — | Caminho completo da imagem de referência |
| `confidence` | `0.80` | Nível de confiança (0.80 a 1.00) |
| `limit` | `10` | Máximo de tentativas |
| `time_ms` | `1000` | Tempo entre tentativas (ms) |
| `offset` | `"center"` | Ponto de clique: left, right, top, bottom, center |
| `offset_x` | `0` | Deslocamento X adicional (pixels) |
| `offset_y` | `0` | Deslocamento Y adicional (pixels) |
| `match_index` | `0` | Índice da ocorrência (0-based) |
| `highlight_ms` | `0` | Tempo para mostrar borda vermelha (ms) |

## Retorno

### find_image / click_image

```python
{
    "x": 540,           # Coordenada X do ponto clicado
    "y": 320,           # Coordenada Y do ponto clicado
    "width": 80,        # Largura da região encontrada
    "height": 30,       # Altura da região encontrada
    "confidence": 0.90, # Confiança utilizada
    "attempts": 1,      # Tentativas realizadas
    "elapsed_ms": 245   # Tempo total (ms)
}
```

### find_all_images

```python
[
    {"index": 0, "x": 540, "y": 320, "width": 80, "height": 30},
    {"index": 1, "x": 200, "y": 150, "width": 80, "height": 30},
]
```

## Offset

| Valor | Descrição |
|-------|-----------|
| `"center"` | Centro da imagem (padrão) |
| `"left"` | Esquerda da imagem |
| `"right"` | Direita da imagem |
| `"top"` | Topo da imagem |
| `"bottom"` | Base da imagem |

## Exemplos Práticos

### Clicar em botão com confiança alta

```python
desktop.click_image("C:/imgs/botao_confirmar.png", confidence=0.95)
```

### Clicar à direita do elemento (ex: dropdown)

```python
desktop.click_image("C:/imgs/seta_dropdown.png", offset="right", offset_x=5)
```

### Encontrar e processar todas as ocorrências

```python
icons = desktop.find_all_images("C:/imgs/icone_lixeira.png")
print(f"Encontrados {len(icons)} itens para deletar")

for icon in icons:
    desktop.click_image("C:/imgs/icone_lixeira.png", match_index=icon["index"])
    time.sleep(0.5)
```

### Debug com highlight

```python
# Mostra borda vermelha por 2 segundos onde encontrou
desktop.click_image("C:/imgs/botao.png", highlight_ms=2000)
```

### Retry com mais tentativas

```python
# Elemento demora para aparecer
desktop.click_image("C:/imgs/botao.png", limit=30, time_ms=2000)
```
