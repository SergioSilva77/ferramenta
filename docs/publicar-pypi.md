# Como Publicar no PyPI

Tutorial rápido para publicar o pacote `rpaflow` no PyPI.

## Pré-requisitos

```bash
pip install build twine
```

## Passo 1: Atualizar Versão

Edite `pyproject.toml` e incremente a versão:

```toml
[project]
name = "rpaflow"
version = "0.1.3"  # ← mudar aqui
```

## Passo 2: Limpar Builds Antigos

```bash
Remove-Item -Recurse -Force dist\*, build\*, *.egg-info
```

## Passo 3: Build

```bash
python -m build
```

Gera dois arquivos em `dist/`:
- `rpaflow-0.1.3-py3-none-any.whl` (wheel)
- `rpaflow-0.1.3.tar.gz` (source)

## Passo 4: Verificar

```bash
python -m twine check dist/*
```

Deve mostrar `PASSED` para ambos.

## Passo 5: Upload

### Usando token do arquivo .env

```bash
# Ler token do .env e fazer upload
$token = (Get-Content .env | Where-Object { $_ -match "^token=" }) -replace "token=",""
python -m twine upload dist/* --username __token__ --password $token
```

### Usando token direto

```bash
python -m twine upload dist/* --username __token__ --password "pypi-SEU_TOKEN_AQUI"
```

## Passo 6: Verificar no PyPI

Acesse: https://pypi.org/project/rpaflow/

## Criar Token no PyPI

1. Acesse https://pypi.org/manage/account/token/
2. Clique em "Add API token"
3. Nome: `rpaflow-upload`
4. Escopo: "Entire account" (ou específico para o projeto)
5. Copie o token (começa com `pypi-`)

## Estrutura do Projeto

```
ferramentarpa/
├── src/
│   └── rpaflow/
│       ├── __init__.py
│       ├── browser.py
│       ├── sql.py
│       └── ...
├── pyproject.toml
├── README.md
├── LICENSE
└── .env              # ← token do PyPI (não commitar!)
```

## Regras Importantes

- **Versão**: Sempre incrementar antes de publicar (PyPI não aceita sobrescrever)
- **README.md**: É exibido na página do PyPI — mantenha atualizado
- **.env**: NUNCA commite o token no Git
- **dist/**: Limpe builds antigos antes de gerar novos
