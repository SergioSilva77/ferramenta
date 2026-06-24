# Excel

Leitura e escrita de planilhas Excel (.xlsx).

## Instalar

```bash
pip install rpaflow[excel]
```

## Exemplo

```python
from rpaflow.excel import Excel

# Abrir planilha
planilha = Excel("dados.xlsx")
planilha.open()

# Ler dados
dados = planilha.read("Planilha1", range="A1:D10")

# Escrever dados
planilha.write("Planilha1", range="A1", values=[
    ["Nome", "Idade", "Cidade"],
    ["Ana", 25, "São Paulo"],
    ["João", 30, "Rio"]
])

# Salvar
planilha.save("saida.xlsx")

# Fechar
planilha.close()
```

## Métodos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `open()` | — | Abre planilha |
| `read()` | sheet, range (opcional) | Lê dados |
| `write()` | sheet, range, values | Escreve dados |
| `save()` | filepath | Salva arquivo |
| `close()` | — | Fecha planilha |
