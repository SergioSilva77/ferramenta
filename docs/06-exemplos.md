# 06 - Exemplos

## Exemplo 1: Hello World

```rpa
# hello.rpa
log --msg="Olá, mundo!"
```

## Exemplo 2: Automação Web com Lógica

```rpa
# web_scraping.rpa
log --msg="Iniciando automação"

startBrowser --url="https://site.com" --type="playwright"

$titulo = getText --selector="h1"
log --msg="Título: " + $titulo

click --selector="#botao-enviar"
screenshot --path="captura.png"

if $titulo == "Site Atualizado"
    log --msg="Site ok!"
elseif $titulo == ""
    log --msg="Falha ao obter título!"
else
    log --msg="Título inesperado: " + $titulo
endif

closeBrowser
log --msg="Fim!"
```

## Exemplo 3: Loop e Condição

```rpa
# loop.rpa
for $i = 1 to 5
    log --msg="Iteração: " + $i
next

$contador = 0
while $contador < 3
    log --msg="Contador: " + $contador
    $contador = $contador + 1
endwhile

if $contador == 3
    log --msg="Loop concluído!"
endif
```

## Exemplo 4: Try/Catch

```rpa
# erro.rpa
try
    startBrowser --url="https://site-invalido.com" --type="playwright"
    $texto = getText --selector="#conteudo"
catch $erro
    log --msg="Erro: " + $erro
finally
    closeBrowser
endtry
```
