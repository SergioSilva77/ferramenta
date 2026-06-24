# 01 - Visão Geral

## O que é MinhaRPA

Linguagem de domínio específico (DSL) criada com ANTLR4 para automação de processos robóticos (RPA). Arquitetura modular com sistema de plugins — o core da linguagem é fixo, os comandos RPA são injetados via plugins conforme necessidade.

## Objetivo

Eliminar a necessidade de escrever boilerplate toda vez que um novo bot RPA é necessário. Em vez de 200+ linhas de Python/C#:

```rpa
startBrowser --url="https://site.com" --type="playwright"
click --selector="#login"
typeText --selector="#user" --text="admin"
openExcel --path="dados.xlsx"
log --msg="Automação finalizada!"
```

## Por que criar uma linguagem própria?

- **Reutilização**: Comandos RPA prontos, sem reimplementar toda vez
- **Legibilidade**: Scripts mais claros que código Python/C#
- **Extensibilidade**: Novos comandos = novo plugin, sem tocar no core
- **Manutenção**: Alterações no plugin afetam todos os bots que usam aquele comando
