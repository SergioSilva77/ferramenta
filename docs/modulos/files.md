# Files

Operações com arquivos e diretórios. Inspirado nos métodos estáticos de `System.IO.File` e `System.IO.Path` do C#.

## Instalar

```bash
pip install rpaflow[files]
```

## Como Usar (Passo a Passo)

```python
from rpaflow.files import Files

files = Files()

# ====== LEITURA/ESCRITA TEXTO ======
conteudo = files.read_text("arquivo.txt", encoding="utf-8")
files.write_text("arquivo.txt", "conteúdo", encoding="utf-8")
files.append_text("arquivo.txt", "novo conteúdo", encoding="utf-8")

# ====== LEITURA/ESCRITA LINHAS ======
linhas = files.read_lines("arquivo.txt")                    # ["linha1", "linha2"]
files.write_lines("arquivo.txt", ["linha1", "linha2"])
files.append_lines("arquivo.txt", ["linha3", "linha4"])

# ====== LEITURA/ESCRITA BYTES ======
dados = files.read_bytes("arquivo.bin")
files.write_bytes("arquivo.bin", b"\x00\x01\x02")
files.append_bytes("arquivo.bin", b"\x03\x04")

# ====== OPERAÇÕES ======
files.exists("arquivo.txt")                                 # True/False
files.copy("origem.txt", "destino.txt")
files.move("antigo.txt", "novo.txt")
files.delete("lixo.txt")
files.replace("novo.txt", "antigo.txt", "backup.txt")

# ====== INFORMAÇÕES ======
files.get_creation_time("arquivo.txt")                      # datetime
files.get_last_write_time("arquivo.txt")                    # datetime
files.get_size("arquivo.txt")                               # bytes

# ====== PATH ======
files.get_filename("C:/pasta/arquivo.txt")                  # "arquivo.txt"
files.get_filename_without_ext("C:/pasta/arquivo.txt")      # "arquivo"
files.get_extension("arquivo.txt")                          # ".txt"
files.get_directory("C:/pasta/arquivo.txt")                 # "C:/pasta"
files.combine("C:/pasta", "arquivo.txt")                    # "C:/pasta/arquivo.txt"
files.get_full_path("arquivo.txt")                          # "C:/atual/arquivo.txt"
files.is_rooted("C:/pasta/arquivo.txt")                     # True
files.has_extension("arquivo.txt")                          # True
files.change_extension("arquivo.txt", ".log")               # "arquivo.log"
files.get_temp_path()                                       # "C:/Users/.../Temp"
files.get_base_directory()                                  # Diretório atual

# ====== DIRECTORY ======
files.dir_create("C:/nova/pasta")
files.dir_delete("C:/pasta/antiga", recursive=True)
files.dir_exists("C:/pasta")                                # True/False
files.dir_move("C:/origem", "C:/destino")
files.dir_get_files("C:/pasta")                             # ["arquivo1.txt", ...]
files.dir_get_dirs("C:/pasta")                              # ["subpasta1", ...]
files.dir_get_all_files("C:/pasta")                         # Lista recursiva

# ====== FILE INFO ======
files.get_size("arquivo.txt")                               # 1234 (bytes)
files.is_readonly("arquivo.txt")                            # True/False
files.set_readonly("arquivo.txt", readonly=True)
```

## Métodos

### Leitura/Escrita Texto

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `read_text()` | filepath, encoding | Lê conteúdo como texto |
| `write_text()` | filepath, content, encoding | Escreve texto no arquivo |
| `append_text()` | filepath, content, encoding | Adiciona texto ao final |
| `read_lines()` | filepath, encoding | Lê linhas como lista |
| `write_lines()` | filepath, lines, encoding | Escreve lista de linhas |
| `append_lines()` | filepath, lines, encoding | Adiciona linhas ao final |

### Leitura/Escrita Bytes

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `read_bytes()` | filepath | Lê conteúdo como bytes |
| `write_bytes()` | filepath, data | Escreve bytes no arquivo |
| `append_bytes()` | filepath, data | Adiciona bytes ao final |

### Operações de Arquivo

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `exists()` | filepath | Verifica se arquivo existe |
| `copy()` | src, dst, overwrite | Copia arquivo |
| `move()` | src, dst, overwrite | Move arquivo |
| `delete()` | filepath | Deleta arquivo |
| `replace()` | source, destination, backup | Substitui arquivo |

### Informações do Arquivo

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get_creation_time()` | filepath | Data de criação |
| `get_last_write_time()` | filepath | Data da última escrita |
| `get_last_access_time()` | filepath | Data do último acesso |
| `get_attributes()` | filepath | Atributos do arquivo |
| `set_attributes()` | filepath, readonly, hidden | Define atributos |
| `create_symlink()` | link, target | Cria link simbólico |

### Path

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get_filename()` | filepath | Nome do arquivo com extensão |
| `get_filename_without_ext()` | filepath | Nome sem extensão |
| `get_extension()` | filepath | Extensão do arquivo |
| `get_directory()` | filepath | Diretório pai |
| `combine()` | *paths | Junta caminhos |
| `join()` | *paths | Junta caminhos (os.path.join) |
| `get_full_path()` | filepath | Caminho absoluto |
| `get_relative_path()` | base, filepath | Caminho relativo |
| `is_rooted()` | filepath | Verifica se é absoluto |
| `is_fully_qualified()` | filepath | Verifica se é completo |
| `has_extension()` | filepath | Verifica se tem extensão |
| `change_extension()` | filepath, new_ext | Troca extensão |
| `trim_separator()` | filepath | Remove separador final |
| `ends_with_separator()` | filepath | Verifica se termina com separador |
| `get_temp_path()` | — | Diretório temporário |
| `get_temp_file()` | — | Arquivo temporário |
| `get_random_filename()` | — | Nome aleatório |
| `get_invalid_filename_chars()` | — | Caracteres inválidos no nome |
| `get_invalid_path_chars()` | — | Caracteres inválidos no caminho |
| `get_path_root()` | filepath | Raiz do caminho |
| `get_base_directory()` | — | Diretório base (AppDomain.BaseDirectory) |

### Directory

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `dir_create()` | path | Cria diretório |
| `dir_delete()` | path, recursive | Deleta diretório |
| `dir_exists()` | path | Verifica se diretório existe |
| `dir_move()` | src, dst | Move diretório |
| `dir_get_files()` | path, pattern | Lista arquivos |
| `dir_get_dirs()` | path | Lista subdiretórios |
| `dir_get_all_files()` | path, pattern | Lista todos recursivamente |

### File Info

| Método | Parâmetros | Descrição |
|--------|-----------|-----------|
| `get_size()` | filepath | Tamanho em bytes |
| `is_readonly()` | filepath | Verifica se é somente leitura |
| `set_readonly()` | filepath, readonly | Define somente leitura |

### Parâmetros Comuns

| Parâmetro | Padrão | Descrição |
|-----------|--------|-----------|
| `encoding` | `"utf-8"` | Encoding do texto |
| `overwrite` | `False` | Sobrescrever se existir |
| `recursive` | `False` | Operação recursiva |
| `pattern` | `"*"` | Padrão de busca |

## Exemplos Práticos

### Ler e processar arquivo CSV

```python
from rpaflow.files import Files

files = Files()

# Ler CSV como linhas
linhas = files.read_lines("dados.csv", encoding="latin-1")

for linha in linhas:
    campos = linha.split(";")
    print(f"Nome: {campos[0]}, Idade: {campos[1]}")
```

### Backup de arquivo

```python
from rpaflow.files import Files
from datetime import datetime

files = Files()

# Criar backup com data
data = datetime.now().strftime("%Y%m%d_%H%M%S")
backup = f"backup_{data}.txt"
files.copy("arquivo.txt", backup)

print(f"Backup criado: {backup}")
```

### Processar todos os arquivos de uma pasta

```python
from rpaflow.files import Files

files = Files()

# Listar todos os .txt recursivamente
arquivos = files.dir_get_all_files("C:/dados", "*.txt")

for arquivo in arquivos:
    conteudo = files.read_text(arquivo)
    print(f"{arquivo}: {len(conteudo)} caracteres")
```

### Verificar e criar diretório

```python
from rpaflow.files import Files

files = Files()

if not files.dir_exists("C:/logs"):
    files.dir_create("C:/logs")
    print("Diretório criado")

# Escrever log
files.append_text("C:/logs/app.log", "Nova entrada de log\n")
```

### Obter informações do arquivo

```python
from rpaflow.files import Files

files = Files()

info = files.get_attributes("arquivo.txt")
print(f"Tamanho: {info['size']} bytes")
print(f"Criado: {info['created']}")
print(f"Modificado: {info['modified']}")
print(f"Somente leitura: {info['is_readonly']}")
```

### Trabalhar com caminhos

```python
from rpaflow.files import Files

files = Files()

caminho = "C:/pasta/subpasta/arquivo.txt"

print(files.get_filename(caminho))              # "arquivo.txt"
print(files.get_filename_without_ext(caminho))  # "arquivo"
print(files.get_extension(caminho))             # ".txt"
print(files.get_directory(caminho))             # "C:/pasta/subpasta"
print(files.get_path_root(caminho))             # "C:\"
print(files.is_rooted(caminho))                 # True

# Combinar caminhos
novo = files.combine("C:/pasta", "subpasta", "arquivo.txt")
print(novo)  # "C:/pasta/subpasta/arquivo.txt"

# Trocar extensão
log = files.change_extension(caminho, ".log")
print(log)  # "C:/pasta/subpasta/arquivo.log"
```
