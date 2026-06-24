# 04 - Constructos da Linguagem

## Estruturas de Controle

| Constructo | Sintaxe |
|------------|---------|
| If/Elseif/Else | `if $x == 1` ... `elseif $x == 2` ... `else` ... `endif` |
| While | `while $i < 10` ... `endwhile` |
| For | `for $i = 0 to 10` ... `next` / `for $i = 0 to 10 step 2` ... `next` |
| Try/Catch | `try` ... `catch $erro` ... `finally` ... `endtry` |
| Funções | `def minhaFuncao($a, $b)` ... `enddef` |

## Variáveis

```rpa
$nome = "João"
$idade = 25
$ativo = true
$lista = [1, 2, 3]
$mapa = {"chave": "valor"}
```

## Expressões

- Aritméticas: `+`, `-`, `*`, `/`
- Comparação: `==`, `!=`, `>`, `<`, `>=`, `<=`
- Lógicas: `and`, `or`, `!`
- Concatenação: `"texto " + $variavel`

## Comandos Genéricos

```rpa
# Comando com flags
startBrowser --url="https://site.com" --type="playwright"

# Comando com posicionais
log "Mensagem aqui"

# Comando sem argumentos
closeBrowser
```
