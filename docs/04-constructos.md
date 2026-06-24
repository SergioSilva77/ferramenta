# 04 - Constructos da Linguagem

## Estruturas de Controle

| Constructo | Sintaxe |
|------------|---------|
| If/Elseif/Else | `if $x == 1` ... `elseif $x == 2` ... `else` ... `endif` |
| While | `while $i < 10` ... `endwhile` |
| For | `for $i = 0 to 10` ... `next` / `for $i = 0 to 10 step 2` ... `next` |
| Foreach | `foreach $item in $lista` ... `endforeach` |
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

## Arrays

```rpa
# Criar array
$frutas = ["maçã", "banana", "laranja"]
$numeros = [10, 20, 30, 40, 50]

# Acessar por índice (começa em 0)
log --msg=$frutas[0]        # maçã
log --msg=$frutas[2]        # laranja

# Modificar valor
$frutas[1] = "uva"

# Percorrer com foreach
foreach $fruta in $frutas
    log --msg="Fruta: " + $fruta
endforeach
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
