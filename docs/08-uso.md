# 08 - Como Usar

## Criando um Script

Crie um arquivo com extensão `.rpa` em qualquer editor de texto:

```rpa
# meu_bot.rpa

log --msg="Iniciando automação..."

startBrowser --url="https://site.com" --type="playwright"

$titulo = getText --selector="h1"
log --msg="Título: " + $titulo

if $titulo == "Site OK"
    log --msg="Sucesso!"
else
    log --msg="Falha!"
endif

closeBrowser
log --msg="Fim!"
```

### Regras

- Extensão: `.rpa`
- Comentários: `#` no início da linha
- Variáveis: iniciam com `$` (ex: `$nome`, `$contador`)
- Strings: entre aspas duplas `"texto"`
- Comandos: `nomeComando --flag=valor` ou `nomeComando "posicional"`

---

## Executando

### Opção 1: Comando CLI

```bash
# Instale as dependências (uma única vez)
pip install -r requirements.txt

# Execute qualquer script
python main.py meu_bot.rpa

# Ou com alias (depois de instalar globalmente)
rpa execute meu_bot.rpa
```

### Opção 2: Direto com Python

```bash
python main.py examples/hello.rpa
```

### Opção 3: Modo interativo (futuro)

```bash
rpa interactive
> log --msg="Hello!"
[LOG] Hello!
> $x = 10
> log --msg=$x
[LOG] 10
>
```

---

## Saída Esperada

```
$ python main.py meu_bot.rpa

  MinhaRPA v1.0
  Executando: meu_bot.rpa

  [LOG] Iniciando automação...
  [LOG] Título: Bem-vindo
  [LOG] Sucesso!
  [LOG] Fim!

  Script finalizado (1.2s)
```

---

## Estrutura de Comandos

### Flags (`--chave=valor`)

```rpa
startBrowser --url="https://site.com" --type="playwright"
click --selector="#btn"
typeText --selector="#input" --text="Olá"
screenshot --path="captura.png"
```

### Posicionais

```rpa
log "Mensagem aqui"
wait 3
```

### Sem argumentos

```rpa
closeBrowser
```

---

## Combinando Comandos com Lógica

```rpa
# Ler dados e decidir o que fazer
$usuario = getText --selector="#user"

if $usuario == "admin"
    log --msg="Acesso liberado"
    startBrowser --url="https://admin.com" --type="playwright"
else
    log --msg="Acesso negado"
endif
```

---

## Loop para Processar Vários Itens

```rpa
# Processar 10 páginas
for $i = 1 to 10
    $url = "https://site.com/pagina/" + $i
    startBrowser --url=$url --type="playwright"
    $dados = getText --selector="#conteudo"
    log --msg="Página " + $i + ": " + $dados
    closeBrowser
next
```

---

## Tratamento de Erros

```rpa
try
    startBrowser --url="https://site.com" --type="playwright"
    click --selector="#inexistente"
catch $erro
    log --msg="Erro: " + $erro
finally
    closeBrowser
endtry
```
