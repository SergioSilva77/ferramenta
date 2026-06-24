# Files

Operações com arquivos e diretórios.

## Instalar

```bash
pip install rpaflow[files]
```

## Exemplo

```python
from rpaflow.files import Files

files = Files()

# Ler
conteudo = files.read("arquivo.txt")

# Escrever
files.write("saida.txt", "conteúdo do arquivo")

# Copiar
files.copy("origem.txt", "destino.txt")

# Mover
files.move("antigo.txt", "pasta/novo.txt")

# Deletar
files.delete("lixo.txt")
```

## Métodos

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `read()` | filepath | Lê conteúdo |
| `write()` | filepath, content | Escreve conteúdo |
| `copy()` | src, dst | Copia arquivo |
| `move()` | src, dst | Move arquivo |
| `delete()` | filepath | Deleta arquivo |
